process.env.NODE_TLS_REJECT_UNAUTHORIZED = 0;
var fs = require('fs');
var URL = require('url');
if(!Array.prototype.includes) Array.prototype.includes = function includes(find) {
	var length = this.length;
	for(var i=0; i<length; i++)
		if(this[i] == find) return true;
	return false;
};
function print() {
	return console.log.apply(this, Array.prototype.slice.call(arguments).concat(['\r']));
}
if((process.argv[2] || '').match(/^[%]\d$/)) process.argv[2] = '';
if((process.argv[3] || '').match(/^[%]\d$/)) process.argv[3] = '';
if((process.argv[4] || '').match(/^[%]\d$/)) process.argv[4] = '';
var url = process.argv[2] || process.exit(102);
url = url.replace(/^["]/, '').replace(/["]$/, '');
var fn = process.argv[3] || process.exit(102);
var trd = Number(process.argv[4] || process.exit(103)) || process.exit(103);
fn = fn.replace(/^["]/, '').replace(/["]$/, '');
var intpath = require('path');
var parsed = {
	dir: intpath.dirname(fn),
	ext: intpath.extname(fn),
	name: intpath.basename(fn).slice(0, intpath.basename(fn).length - intpath.extname(fn).length),
};
if(fs.existsSync(fn)) {
	if(Number(process.argv[6]) == 0) return process.exit(104);
	else if(Number(process.argv[6]) == 1) fs.unlinkSync(fn);
	else if(Number(process.argv[6]) == 2) fn = parsed.dir.replace(/\\$/, '') + '\\' + parsed.name + '-' + Math.floor(Math.random() * 10000000000000000) + parsed.ext;
	print('MODIFIEDFILENAME', fn);
}
if(fn[fn.length - 1] == '.') {
	fn = fn.replace(/[.]$/, '_');
	print('MODIFIEDFILENAME', fn);
}
var continuedownload = false;
if(process.argv[7] == 1) {
	if(trd <= 1 && fs.existsSync(fn + '.part.tmp')) {
		continuedownload = true;
	} else if(fs.existsSync(fn + '.part_' + trd + '.tmp') && !fs.existsSync(fn + '.part_' + (trd + 1) + '.tmp')) {
		continuedownload = true;
	} else {
		if(trd <= 1 && fs.existsSync(fn + '.part.tmp')) fs.unlinkSync(fn + '.part.tmp');
		for(var i=1; i<=trd+25; i++)
			if(fs.existsSync(fn + '.part_' + i + '.tmp'))
				fs.unlinkSync(fn + '.part_' + i + '.tmp');
	}
} else {
	if(trd <= 1 && fs.existsSync(fn + '.part.tmp')) fs.unlinkSync(fn + '.part.tmp');
	for(var i=1; i<=trd+25; i++)
		if(fs.existsSync(fn + '.part_' + i + '.tmp'))
			fs.unlinkSync(fn + '.part_' + i + '.tmp');
}
var http = require(url.slice(0, 6) == 'https:' ? 'https' : 'http');
print('STATUS', 'CHECKREDIRECT');
var userAgent = 'Mozilla/5.0 (Windows NT 6.1; rv:121.0) Gecko/20100101 Firefox/121.0';
checkRedirect(url);
function checkRedirect(url) {
	var parsedURL = URL.parse(url);
	http.request({
		host: parsedURL.host,
		path: parsedURL.path,
		headers: { 'User-Agent': userAgent },
		method: 'HEAD',
	}, function(res) {
		if(res.headers.location && [301, 302, 303, 307, 308].includes(res.statusCode || 0))
			return checkRedirect(res.headers.location.trim().replace(/^["]/, '').replace(/["]$/, ''));
		print('REALADDR', url);
		startDownload(url);
	}).end();
}
function startDownload(url) {
	if(trd > 1) print('STATUS', 'CHECKFILE');
	var parsedURL = URL.parse(url);
	http.request({
		host: parsedURL.host,
		path: parsedURL.path,
		headers: { 'User-Agent': userAgent },
		method: 'HEAD',
	}, function(res) {
		if((res.statusCode + '')[0] != 2) {
			print('STATUSCODE', res.statusCode + '');
			return process.exit(108);
		}
		var total = Number(res.headers['content-length']);
		if(trd > 1) {
			if(res.headers['accept-ranges'] != 'bytes') return process.exit(106);
			if(!total) return process.exit(107);
		} else if(!total) {
			total = 0;
		}
		if(continuedownload && (res.headers['accept-ranges'] != 'bytes' || !total)) {
			print('STATUS', 'UNABLETOCONTINUE');
			continuedownload = false;
			if(trd <= 1 && fs.existsSync(fn + '.part.tmp')) fs.unlinkSync(fn + '.part.tmp');
			for(var i=1; i<=trd+25; i++)
				if(fs.existsSync(fn + '.part_' + i + '.tmp'))
					fs.unlinkSync(fn + '.part_' + i + '.tmp');
		} else if(res.headers['accept-ranges'] != 'bytes' || !total) {
			print('STATUS', 'RESUMEUNSUPPORTED');
		}
		var comp = 0;
		var downloader = [];
		var totals = [];
		var unit = Math.floor(total / trd);
		var range = 0;
		function get(id, callback) {
			var startRange, endRange;
			var headers = { 'User-Agent': userAgent };
			if(trd > 1) {
				if(continuedownload && fs.existsSync(fn + '.part_' + id + '.tmp')) {
					downloader[id] = fs.statSync(fn + '.part_' + id + '.tmp').size;
					startRange = range + downloader[id];
					endRange = range + unit;
					if(endRange - startRange < 0) {
						ready.push(id);
						totals[id] = downloader[id];
						return callback('NONEEDTODOWNLOAD');
					}
					totals[id] = (endRange >= total ? (total - 1) : endRange) - range + 1;
					headers.Range = 'bytes=' + startRange + '-' + endRange;
				} else {
					headers.Range = 'bytes=' + range + '-' + (range + unit);
				}
			} else if(continuedownload) {
				downloader[id] = fs.statSync(fn + '.part.tmp').size;
				totals[id] = total;
				startRange = downloader[id];
				endRange = total + 1;
				if(endRange - startRange < 0) {
					print('STATUS', 'COMPLETE');
					process.exit(0);
				}
				headers.Range = 'bytes=' + startRange + '-' + endRange;
			}
			return http.get({
				host: parsedURL.host,
				path: parsedURL.path,
				headers: headers,
			}, function(response) {
				return callback(response);
			}).end();
		}
		var ready = [];
		print('STATUS', 'DOWNLOADING');
		(function startThreads(i) {
			if(i > trd) return;
			var id = i;
			get(id, function(response) {
				if(response == 'NONEEDTODOWNLOAD') {
					range += totals[id];
					comp++;
					return startThreads(i + 1);
				}
				if((response.statusCode + '')[0] != 2) {
					print('STATUSCODE', response.statusCode + '');
					return process.exit(108);
				}
				ready.push(i);
				if(!downloader[id]) downloader[id] = 0;
				if(!totals[id]) totals[id] = Number(response.headers['content-length'] || 0);
				range += totals[id];
				response.on('error', function() {});
				response.on('data', function(chunk) {
					downloader[id] += chunk.length;
					fs.appendFileSync(fn + (trd <= 1 ? '.part.tmp' : ('.part_' + id + '.tmp')), chunk);
				});
				response.on('end', function() {
					comp++;
				});
				return setTimeout(function() {
					return startThreads(i + 1);
				}, 100);
			});
		})(1);
		var statusReporter = setInterval(function() {
			try {
				var totalbytes = '';
				var prt = '';
				var psum = 0;
				var dsum = 0;
				for(var di=1; di<=trd; di++) {
					var dn = downloader[di];
					if(dn === undefined) {
						print('DATA', di + ',-1,0,0');
						continue;
					}
					var pc;
					if(totals[di] <= 0) pc = -1;
					else pc = (dn / totals[di]) * 100;
					psum += pc;
					dsum += (dn || 0);
					if(ready.includes(di)) print('DATA', di + ',' + (total == 0 || Math.floor(pc) > 100.0 ? '-1' : Math.floor(pc)) + ',' + totals[di] + ',' + dn);
					else print('DATA', di + ',-1,0,0');
				}
				print('TOTAL', (!total ? '-1' : total) + ',' + dsum + ',' + (total == 0 || psum < 0 ? '-1' : (Math.floor((psum / (100 * trd)) * 100) || '-1')));
				if(comp >= trd) {
					clearInterval(statusReporter);
					if(trd > 1) {
						print('STATUS', 'MERGING');
						var s = 'COPY /B ';
						for(var i=1; i<=trd; i++) s += '"' + fn + '.part_' + i + '.tmp"+';
						s = s.replace(/[+]$/, '');
						s += ' "' + fn + '"';
						require('child_process').exec(s, function() {
							if(Number(process.argv[5]) == 0)
								for(var i=1; i<=trd; i++)
									fs.unlinkSync(fn + '.part_' + i + '.tmp');
							print('STATUS', 'COMPLETE');
							process.exit(0);
						});
					} else {
						fs.renameSync(fn + '.part.tmp', fn);
						print('STATUS', 'COMPLETE');
						process.exit(0);
					}
				}
			} catch (e) {}
		}, 100);
	}).end();
}

setInterval(function() {}, 987654321);

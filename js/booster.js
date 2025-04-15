process.env.NODE_TLS_REJECT_UNAUTHORIZED = 0;
var iconv = require('./iconv');
var fs = require('fs');
var URL = require('url');
if(!Array.prototype.includes) Array.prototype.includes = function includes(find) {
	var length = this.length;
	for(var i=0; i<length; i++)
		if(this[i] == find) return true;
	return false;
};
if(!Object.assign)
	Object.assign = function assign(target, varArgs) {
		if(!target) throw new TypeError('Cannot convert undefined or null to object');
		var to = Object(target);
		for(var index=1; index<arguments.length; index++) {
			var nextSource = arguments[index];
			if(nextSource != null) {
				for(var nextKey in nextSource)
					if(Object.prototype.hasOwnProperty.call(nextSource, nextKey))
						to[nextKey] = nextSource[nextKey];
			}
		}
		return to;
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
var http = require(url.slice(0, 6) == 'https:' ? 'https' : 'http');
var rawHeaders = Buffer(((process.argv[14] == '-' ? '' : process.argv[14]) || ''), 'base64').toString().split('\n');
var rawSessionHeaders = Buffer(((process.argv[15] == '-' ? '' : process.argv[15]) || ''), 'base64').toString().split('\n');
var headers = {}, sessionHeaders = {};
if(rawHeaders.length) rawHeaders.forEach(function(item) {
	if(item.indexOf(': ') < 0) return;
	headers[item.slice(0, item.indexOf(': '))] = item.slice(item.indexOf(': ') + 2);
});
if(rawSessionHeaders.length) rawSessionHeaders.forEach(function(item) {
	if(item.indexOf(': ') < 0) return;
	sessionHeaders[item.slice(0, item.indexOf(': '))] = item.slice(item.indexOf(': ') + 2);
});
Object.assign(headers, sessionHeaders);
function safeFilename(filename) {
	return iconv.decode(iconv.encode(filename, 'cp949'), 'cp949').replace(/\?/g, '_').replace(/\*/g, '_').replace(/\\/g, '_').replace(/\//g, '_').replace(/\:/g, '_').replace(/\"/g, '_').replace(/\|/g, '_').replace(/\</g, '_').replace(/\>/g, '_');
}
if(process.argv[8] == 1) {
	startDownload(url);
} else {
	print('STATUS', 'CHECKREDIRECT');
	checkRedirect(url);
}
function checkRedirect(url) {
	var parsedURL = URL.parse(url);
	http.request({
		host: parsedURL.host,
		path: parsedURL.path,
		headers: headers,
		method: (process.argv[9] == 1) ? 'GET' : 'HEAD',
	}, function(res) {
		if(process.argv[9] == 1) try {
			res.connection.end();
			res.connection.destroy();
		} catch(e) {}
		if(res.headers.location && [301, 302, 303, 307, 308].includes(res.statusCode || 0)) {
			return checkRedirect(res.headers.location.trim().replace(/^["]/, '').replace(/["]$/, ''));
		}
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
		headers: headers,
		method: (process.argv[9] == 1) ? 'GET' : 'HEAD',
	}, function(res) {
		if(process.argv[10] == 1 ? (!(Number((res.statusCode + '')[0]) <= 3)) : ((res.statusCode + '')[0] != 2)) {
			print('STATUSCODE', res.statusCode + '');
			return process.exit(108);
		}
		
		var lastModified = new Date(res.headers['last-modified']);
		var parsed = {
			dir: intpath.dirname(fn),
			base: intpath.basename(fn),
		};
		var sf = safeFilename(parsed.base);
		var fnupd = false;
		if(sf != parsed.base) {
			parsed.base = sf;
			fnupd = true;
		}
		if(fnupd) {
			fn = parsed.dir.replace(/\\$/, '') + '\\' + sf;
			print('MODIFIEDFILENAME', iconv.encode(fn, 'cp949').toString('base64'));
		}
		var disposition = res.headers['content-disposition'];
		if(process.argv[11] == 1 && disposition) {
			var filename = parsed.base;
			/* https://stackoverflow.com/questions/40939380/how-to-get-file-name-from-content-disposition */
			var utf8FilenameRegex = /filename\*=UTF-8''([\w%\-\.]+)(?:; ?|$)/i;
			var asciiFilenameRegex = /^filename=(["']?)(.*?[^\\])\1(?:; ?|$)/i;
			if (utf8FilenameRegex.test(disposition)) {
				filename = decodeURIComponent(utf8FilenameRegex.exec(disposition)[1]);
			} else {
				var filenameStart = disposition.toLowerCase().indexOf('filename=');
				if (filenameStart >= 0) {
					var partialDisposition = disposition.slice(filenameStart);
					var matches = asciiFilenameRegex.exec(partialDisposition);
					if (matches != null && matches[2])
						filename = matches[2];
				}
			}
			filename = filename.trim();
			if(filename && filename != parsed.base) {
				filename = safeFilename(filename);
				fn = parsed.dir.replace(/\\$/, '') + '\\' + filename;
				print('MODIFIEDFILENAME', iconv.encode(fn, 'cp949').toString('base64'));
			}
		}
		parsed.dir = intpath.dirname(fn);
		parsed.ext = intpath.extname(fn);
		parsed.base = intpath.basename(fn);
		parsed.name = parsed.base.slice(0, parsed.base.length - parsed.ext.length);
		if(fs.existsSync(fn)) {
			if(Number(process.argv[6]) == 0) return process.exit(104);
			else if(Number(process.argv[6]) == 1) fs.unlinkSync(fn);
			else if(Number(process.argv[6]) == 2) {
				fn = parsed.dir.replace(/\\$/, '') + '\\' + parsed.name + '-' + Math.floor(Math.random() * 10000000000000000) + parsed.ext;
				print('MODIFIEDFILENAME', iconv.encode(fn, 'cp949').toString('base64'));
			}
		}
		if(fn[fn.length - 1] == '.') {
			fn = fn.replace(/[.]$/, '_');
			print('MODIFIEDFILENAME', iconv.encode(fn, 'cp949').toString('base64'));
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
		if(process.argv[9] == 1) try {
			res.connection.end();
			res.connection.destroy();
		} catch(e) {}
		var comp = 0;
		var downloader = [];
		var totals = [];
		var unit = Math.floor(total / trd);
		var range = 0;
		function get(id, callback) {
			var startRange, endRange;
			var reqHeaders = Object.assign({}, headers);
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
					reqHeaders.Range = 'bytes=' + startRange + '-' + endRange;
				} else {
					reqHeaders.Range = 'bytes=' + range + '-' + (range + unit);
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
				reqHeaders.Range = 'bytes=' + startRange + '-' + endRange;
			}
			return http.get({
				host: parsedURL.host,
				path: parsedURL.path,
				headers: reqHeaders,
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
				if(process.argv[10] == 1 ? (!(Number((response.statusCode + '')[0]) <= 3)) : ((response.statusCode + '')[0] != 2)) {
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
				}, Number(process.argv[12]) || 100);
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
						
						var mergedsize = 0;
						var sizeReporter = setInterval(function() {
							fs.stat(fn, function(err, stat) {
								if(!err && stat) print('MERGESIZE', stat.size.toString());
							});
						}, 100);
						require('child_process').exec(s, function() {
							clearInterval(sizeReporter);
							print('MERGESIZE', total);
							if(Number(process.argv[5]) == 0) {
								for(var i=1; i<=trd; i++)
									fs.unlinkSync(fn + '.part_' + i + '.tmp');
									// print('DELETEITEM', fn + '.part_' + i + '.tmp');
							}
							setLastModified(fn, lastModified);
							print('STATUS', 'COMPLETE');
							process.exit(0);
						});
					} else {
						fs.renameSync(fn + '.part.tmp', fn);
						setLastModified(fn, lastModified);
						print('STATUS', 'COMPLETE');
						process.exit(0);
					}
				}
			} catch (e) {}
		}, 100);
	}).end();
}

function setLastModified(fn, lastModified) {
	if(process.argv[13] != 1) return;
	if(lastModified == 'Invalid Date') return;
	// https://www.vbforums.com/showthread.php?704979-How-to-change-file-date
	print('SETMODIFIEDDATE', '1');
}

setInterval(function() {}, 987654321);

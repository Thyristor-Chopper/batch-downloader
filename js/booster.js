process.env.NODE_TLS_REJECT_UNAUTHORIZED = 0;
var iconv = require('./iconvp');
var fs = require('fs');
var URL = require('url');
if(!Array.prototype.includes) Array.prototype.includes = function includes(find) {
	var length = this.length;
	for(var i=0; i<length; i++)
		if(this[i] === find) return true;
	return false;
};
if(!Object.assign) Object.assign = function assign(target, varArgs) {
	var to = Object(target);
	var length = arguments.length;
	for(var idx=1; idx<length; idx++) {
		var nextSource = arguments[idx];
		if(nextSource != null)
			for(var nextKey in nextSource)
				if(Object.prototype.hasOwnProperty.call(nextSource, nextKey))
					to[nextKey] = nextSource[nextKey];
	}
	return to;
};
if(!Buffer.from) Buffer.from = function from(data, encoding) {
	return new Buffer(data, encoding);
};
if(!Buffer.alloc) Buffer.alloc = function alloc(length) {
	return new Buffer(length).fill(0);
};
if(!Buffer.prototype.indexOf) Buffer.prototype.indexOf = function indexOf(val, offset) {
	offset = offset >>> 0;
	var buf = this;
	if(typeof val == 'string') {
		val = new Buffer(val);
	} else if(typeof val == 'number') {
		val = new Buffer([val & 0xFF]);
	} else if(!Buffer.isBuffer(val)) {
		throw TypeError();
	}
	var len = val.length;
	if(!len) return -1;
	for(var i=offset; i<=buf.length-len; i++) {
		var match = true;
		for(var j=0; j<len; j++)
			if(buf[i + j] !== val[j]) {
				match = false;
				break;
			}
		if(match)
			return i;
	}
	return -1;
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
function splitRanges(fileSize, threadCount) {
	var baseSize = Math.floor(fileSize / threadCount);
	var remainder = fileSize % threadCount;
	var ranges = [, ];
	var start = 0;
	for(var i=1; i<=threadCount; i++) {
		var size = baseSize + (i <= remainder ? 1 : 0);
		var end = start + size - 1;
		ranges.push([start, end]);
		start = end + 1;
	}
	return ranges;
}
function intersect(r1, r2) {
	var s = Math.max(r1[0], r2[0]);
	var e = Math.min(r1[1], r2[1]);
	return s <= e ? [s, e] : null;
}
function remapDownloadInfo(downloadInfo, newThreadCount) {
    var totalSize = downloadInfo.downloadRanges[downloadInfo.threads][1] + 1;
    var fileEnd = totalSize - 1;
    var newDownloadRanges = splitRanges(totalSize, newThreadCount);
    var newDownloadedSizes = [, ];
    for(var i=1; i<=newThreadCount; i++) {
        var newRange = newDownloadRanges[i];
        var d = [];
        for(var j=1; j<downloadInfo.downloadRanges.length; j++) {
            var oldDownloaded = downloadInfo.downloadedSizes[j] || [];
			var inter;
            for(var k=0; k<oldDownloaded.length; k++) {
                inter = intersect(oldDownloaded[k], newRange);
                if(inter) d.push(inter);
            }
        }
        newDownloadedSizes.push(d);
    }
    return {
        threads: newThreadCount,
        downloadedSizes: newDownloadedSizes,
        downloadRanges: newDownloadRanges
    };
}
function invertRanges(totalStart, totalEnd, downloaded) {
	var ret = [];
	downloaded.sort(function(l, r) {
		return l[0] - r[0];
	});
	var cur = totalStart;
	for(var i=0; i<downloaded.length; i++) {
		var s = downloaded[i][0], e = downloaded[i][1];
		if(cur < s)
			ret.push([cur, s - 1]);
		if(e + 1 > cur)
			cur = e + 1;
	}
	if(cur <= totalEnd)
		ret.push([cur, totalEnd]);
	return ret;
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
		if(res.headers.location && [301, 302, 303, 307, 308].includes(Number(res.statusCode) || 0)) {
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
		if(sf != parsed.base) parsed.base = sf, fnupd = true;
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
			if(utf8FilenameRegex.test(disposition)) {
				filename = decodeURIComponent(utf8FilenameRegex.exec(disposition)[1]);
			} else {
				var filenameStart = disposition.toLowerCase().indexOf('filename=');
				if(filenameStart >= 0) {
					var partialDisposition = disposition.slice(filenameStart);
					var matches = asciiFilenameRegex.exec(partialDisposition);
					if(matches && matches[2])
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
		if(fs.existsSync(fn + '.part.tmp')) {
			if(process.argv[7] == 1) {
				continuedownload = true;
			} else {
				fs.unlinkSync(fn + '.part.tmp');
			}
		}
		var total = Number(res.headers['content-length']);
		if(trd > 1) {
			if(res.headers['accept-ranges'] != 'bytes') return process.exit(106);
			if(!total) return process.exit(107);
		} else if(!total) {
			total = 0;
		}
		if(res.headers['accept-ranges'] != 'bytes' || !total) {
			if(continuedownload) {
				print('STATUS', 'UNABLETOCONTINUE');
				continuedownload = false;
				if(fs.existsSync(fn + '.part.tmp'))
					fs.unlinkSync(fn + '.part.tmp');
			} else {
				print('STATUS', 'RESUMEUNSUPPORTED');
			}
		}
		if(process.argv[9] == 1) try {
			res.connection.end();
			res.connection.destroy();
		} catch(e) {}
		var downloadInfo = null;
		var fd;
		try {
			fd = fs.openSync(fn + '.part.tmp', 'r+');
		} catch(e) {
			fd = fs.openSync(fn + '.part.tmp', 'w+');
		}
		var prevSize = fs.fstatSync(fd).size;
		if(total) {
			if(prevSize < total) {
				try {
					fs.ftruncateSync(fd, total);
				} catch(e) {
					fs.closeSync(fd);
					if(e.code == 'ENOSPC')
						process.exit(110);
					else
						process.exit(111);
					throw Error();
				}
				downloadInfo = { threads: trd, downloadedSizes: [ , [ [ 0, prevSize - 1 ] ] ] };
			} else if(prevSize > total) {
				var infoBuf = Buffer.alloc(prevSize - total);
				fs.readSync(fd, infoBuf, 0, prevSize - total, total);
				var nulpos = infoBuf.indexOf(0);
				if(nulpos >= 0)
					infoBuf = infoBuf.slice(0, nulpos);
				try {
					downloadInfo = JSON.parse(infoBuf);
				} catch(e) {
					downloadInfo = { threads: trd, downloadedSizes: [, ] };
					for(var i=1; i<=trd; i++)
						downloadInfo.downloadedSizes.push([]);
				}
			} else {
				downloadInfo = { threads: trd, downloadedSizes: [, ] };
				for(var i=1; i<=trd; i++)
					downloadInfo.downloadedSizes.push([]);
			}
			for(var i=1; i<downloadInfo.threads; i++) {
				if(!downloadInfo.downloadedSizes[i])
					downloadInfo.downloadedSizes[i] = [];
				else 
					for(var j=downloadInfo.downloadedSizes[i].length-1; j>=0; j--)
						if(downloadInfo.downloadedSizes[i][j][1] - downloadInfo.downloadedSizes[i][j][0] < 0)
							downloadInfo.downloadedSizes[i].splice(j, 1);
			}
			if(!downloadInfo.downloadRanges)
				downloadInfo.downloadRanges = splitRanges(total, downloadInfo.threads);
			if(trd != downloadInfo.threads)
				downloadInfo = remapDownloadInfo(downloadInfo, trd);
		}
		var comp = 0;
		var downloader = [];
		var totals = [];
		function get(id, callback) {
			var startRange, endRange;
			var reqHeaders = Object.assign({}, headers);
			var ranges = '';
			var dr = [];
			if(downloadInfo && !downloadInfo.downloadedSizes[id])
				downloadInfo.downloadedSizes[id] = [];
			if(continuedownload) {
				downloader[id] = 0;
				for(var i=0; i<downloadInfo.downloadedSizes[id].length; i++)
					downloader[id] += (downloadInfo.downloadedSizes[id][i][1] - downloadInfo.downloadedSizes[id][i][0] + 1);
				startRange = downloadInfo.downloadRanges[id][0];
				endRange = downloadInfo.downloadRanges[id][1];
				if(endRange >= total || id == trd)
					endRange = total - 1;
				if(endRange < startRange) {
					ready.push(id);
					return callback('NONEEDTODOWNLOAD');
				}
				totals[id] = endRange - startRange + 1;
				if(downloader[id] >= totals[id]) {
					ready.push(id);
					return callback('NONEEDTODOWNLOAD');
				}
				var notDownloaded = invertRanges(startRange, endRange, downloadInfo.downloadedSizes[id]);
				for(var i=0; i<notDownloaded.length; i++)
					dr.push(notDownloaded[i][0] + '-' + notDownloaded[i][1]);
				reqHeaders.Range = 'bytes=' + dr.join(',');
			} else if(trd > 1) {
				startRange = downloadInfo.downloadRanges[id][0];
				endRange = downloadInfo.downloadRanges[id][1];
				if(endRange >= total || id == trd)
					endRange = total - 1;
				if(endRange < startRange) {
					ready.push(id);
					return callback('NONEEDTODOWNLOAD');
				}
				var nrange = startRange + '-' + endRange;
				reqHeaders.Range = 'bytes=' + nrange;
				dr.push(nrange);
			} else if(total) {
				dr.push('0-' + (total - 1));
			}
			return http.get({
				host: parsedURL.host,
				path: parsedURL.path,
				headers: reqHeaders,
			}, function(response) {
				if(downloadInfo)
					callback(response, dr, downloadInfo.downloadedSizes[id]);
				else
					callback(response, dr, null);
			}).end();
		}
		var ready = [];
		print('STATUS', 'DOWNLOADING');
		(function startThreads(i) {
			if(i > trd) return;
			var id = i;
			get(id, function(response, dr, downloadedSizes) {
				if(response == 'NONEEDTODOWNLOAD') {
					comp++;
					return startThreads(i + 1);
				}
				if(process.argv[10] == 1 ? (!(Number((response.statusCode + '')[0]) <= 3)) : ((response.statusCode + '')[0] != 2)) {
					print('STATUSCODE', response.statusCode + '');
					fs.closeSync(fd);
					return process.exit(108);
				}
				ready.push(i);
				if(!downloader[id]) downloader[id] = 0;
				if(!totals[id]) totals[id] = Number(response.headers['content-length'] || 0);
				var sizeidx;
				response.on('error', function() {});
				var contentType = response.headers['content-type'];
				if(total && contentType && contentType.toLowerCase().indexOf('multipart/byteranges') == 0) {
					var m = /boundary=([^\s;]+)/i.exec(contentType);
					if(!m) {
						fs.closeSync(fd);
						process.exit(109);
						throw Error();
					}
					var boundary = '--' + m[1];
					var buffer = Buffer.alloc(0);
					var currentPart = null;
					var bytesWrittenInPart = 0;
					var currentPartIndex = 0;
					response.on('data', function(chunk) {
						buffer = Buffer.concat([buffer, chunk]);
						while(buffer.length) {
							if(!currentPart) {
								var bidx = buffer.indexOf(boundary);
								if(bidx == -1) break;
								buffer = buffer.slice(bidx + boundary.length);
								if(buffer.slice(0, 2).toString() == '--') {
									buffer = buffer.slice(2);
									break;
								}
								var headerEnd = buffer.indexOf('\r\n\r\n');
								if(headerEnd == -1) break;
								var headerStr = buffer.slice(0, headerEnd).toString();
								var crMatch = /Content-Range:\s*bytes\s+(\d+)-(\d+)/i.exec(headerStr);
								var start, end;
								if(!crMatch) {
									if(!dr || !dr[currentPartIndex]) {
										fs.closeSync(fd);
										process.exit(109);
										throw Error();
									}
									var rgsplit = dr[currentPartIndex].split('-');
									start = rgsplit[0];
									end = rgsplit[1];
									if(start == undefined || end == undefined) {
										fs.closeSync(fd);
										process.exit(109);
										throw Error();
									}
								} else {
									start = parseInt(crMatch[1], 10);
									end = parseInt(crMatch[2], 10);
								}
								currentPart = { start: start, end: end };
								sizeidx = downloadedSizes.push([start, start - 1]) - 1;
								currentPartIndex++;
								bytesWrittenInPart = 0;
								buffer = buffer.slice(headerEnd + 4);
							}
							var partBytesRemaining = currentPart.end - currentPart.start + 1 - bytesWrittenInPart;
							var writtenBytes;
							if(buffer.length <= partBytesRemaining) {
								fs.writeSync(fd, buffer, 0, buffer.length, currentPart.start + bytesWrittenInPart);
								writtenBytes = buffer.length;
								buffer = Buffer.alloc(0);
							} else {
								fs.writeSync(fd, buffer, 0, partBytesRemaining, currentPart.start + bytesWrittenInPart);
								writtenBytes = partBytesRemaining;
								buffer = buffer.slice(partBytesRemaining);
							}
							downloadedSizes[sizeidx][1] = currentPart.start + bytesWrittenInPart + writtenBytes - 1;
							bytesWrittenInPart += writtenBytes;
							downloader[id] += writtenBytes;
							if(bytesWrittenInPart == currentPart.end - currentPart.start + 1) {
								currentPart = null;
								bytesWrittenInPart = 0;
							}
						}
					});
				} else {
					var pos, start, sizeidx;
					if(total)
						start = parseInt((dr[0] || '0').split('-')[0], 10) || 0;
					else
						start = 0;
					pos = start;
					if(total && downloadedSizes)
						sizeidx = downloadedSizes.push([start, start - 1]) - 1;
					response.on('data', function(chunk) {
						fs.writeSync(fd, chunk, 0, chunk.length, pos);
						pos += chunk.length;
						downloader[id] += chunk.length;
						if(total) {
							if(downloadedSizes)
								downloadedSizes[sizeidx][1] = pos - 1;
						}
					});
				}
				response.on('end', function() {
					comp++;
				});
				return setTimeout(function() {
					return startThreads(i + 1);
				}, Number(process.argv[12]) || 100);
			});
		})(1);
		var infoSaveInterval = 0;
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
				if(infoSaveInterval >= 30) {
					infoSaveInterval = 0;
					var downloadInfoJSON = Buffer.concat([Buffer.from(JSON.stringify(downloadInfo)), Buffer.alloc(1)]);
					try {
						fs.writeSync(fd, downloadInfoJSON, 0, downloadInfoJSON.length, total);
					} catch(e) {
						fs.closeSync(fd);
						if(e.code == 'ENOSPC')
							process.exit(110);
						else
							process.exit(111);
						throw Error();
					}
				} else {
					infoSaveInterval++;
				}
				if(comp >= trd) {
					if(total && dsum < total) {
						fs.closeSync(fd);
						process.exit(1);
						throw Error();
					}
					clearInterval(statusReporter);
					if(total) fs.ftruncateSync(fd, total);
					fs.closeSync(fd);
					fs.renameSync(fn + '.part.tmp', fn);
					setLastModified(lastModified);
					print('STATUS', 'COMPLETE');
				}
			} catch (e) {}
		}, 100);
	}).end();
}
function setLastModified(lastModified) {
	if(process.argv[13] == 0) return;
	if(lastModified == 'Invalid Date') return;
	var dateStr = lastModified.getFullYear() + '-';
	dateStr += (lastModified.getMonth() + 1) + '-';
	dateStr += lastModified.getDate() + ' ';
	dateStr += lastModified.getHours() + ':';
	dateStr += lastModified.getMinutes() + ':';
	dateStr += lastModified.getSeconds();
	print('SETMODIFIEDDATE', dateStr);
}
setInterval(function() {}, 987654321);

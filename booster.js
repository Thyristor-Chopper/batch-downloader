const fs = require('fs');
const URL = require('url');
if(!Array.prototype.includes) Array.prototype.includes = function includes(find) {
    for(var item of this)
        if(item == find) return !0;
    return !1;
};

function timeout(ms) {
    return new Promise((resolve, reject) => {
        setTimeout(() => resolve(0), ms);
    });
}

function print() {
    return console.log.apply(this, Array.prototype.slice.call(arguments).concat(['\r']));
}
if((process.argv[2] || '').match(/^[%]\d$/)) process.argv[2] = '';
if((process.argv[3] || '').match(/^[%]\d$/)) process.argv[3] = '';
if((process.argv[4] || '').match(/^[%]\d$/)) process.argv[4] = '';
var url = process.argv[2] || process.exit(2);
var fn = process.argv[3] || process.exit(2);
var trd = Number(process.argv[4] || process.exit(3)) || process.exit(3);
fn = fn.replace(/^["]/, '').replace(/["]$/, '');
const parsed = require('path').parse(fn);
if(fs.existsSync(fn)) {
    if(Number(process.argv[6]) == 0) return process.exit(4);
    else if(Number(process.argv[6]) == 1) fs.unlinkSync(fn, () => 1);
    else if(Number(process.argv[6]) == 2) fn = parsed.dir.replace(/\\$/, '') + '\\' + parsed.name + '-' + Math.floor(Math.random() * 10000000000000000) + parsed.ext;
    print('MODIFIEDFILENAME', fn);
}
if(fn.endsWith('.')) {
    fn = fn.replace(/[.]$/, '_');
    print('MODIFIEDFILENAME', fn);
}
if(trd <= 1 && fs.existsSync(fn + '.tmp')) fs.unlinkSync(fn + '.tmp');
for(var i = 1; i <= trd; i++)
    if(fs.existsSync(fn + '.part.' + i)) return process.exit(5);
const http = require(url.startsWith('https:') ? 'https' : 'http');
print('STATUS', 'CHECKREDIRECT');
http.get(url.replace(/^["]/, '').replace(/["]$/, ''), res => {
    if(res.headers.location) url = res.headers.location, print('REALADDR', url);
    if(trd > 1) print('STATUS', 'CHECKFILE');
    http.get(url, res => {
        res.setEncoding('base64');
        var total = Number(res.headers['content-length']);
        if(trd > 1) {
            if(res.headers['accept-ranges'] != 'bytes') return process.exit(6);
            if(!total) return process.exit(7);
        } else if(!total) {
            total = 0
        }
        var completed = [], comp = 0;
        var downloader = [];
        var downloads = [];
        var totals = [];
        var tt = [];
        var unit = Math.floor(total / trd);
        var range = 0;
        function get() {
            var headers = {};
            if(trd > 1) headers.Range = 'bytes=' + range + '-' + (range + unit);
            return new Promise((resolve, reject) => {
                http.get({
                    host: URL.parse(url).host,
                    path: URL.parse(url).path,
                    headers,
                }, res => resolve(res));
            });
        }
        var ready = [];
        print('STATUS', 'DOWNLOADING');
        (async function() {
			var i;
            for(i=1, range=0; i<=trd; i++) {
                const response = await get();
                if((response.statusCode + '')[0] != 2) {
                    await timeout(100);
                    continue;
                }
                const id = i;
                ready.push(i);
                downloader[id] = 0;
                downloads[id] = '';
                completed[id] = 0;
                totals[id] = tt[id] = Number(response.headers['content-length'] || 0);
                range += totals[id];
                response.on('error', () => 1);
                response.on('data', chunk => (downloader[id] += chunk.length, fs.appendFileSync(fn + (trd <= 1 ? '.tmp' : ('.part.' + id)), chunk)));
                response.on('end', () => comp++, completed[id] = 1);
                await timeout(100);
            }
        })();
        var statusReporter = setInterval(async () => {
            try {
                var totalbytes = '';
                var prt = '';
                var psum = 0;
                var dsum = 0;
                for(di = 1; di <= trd; di++) {
                    if(di == 'includes') continue;
                    var dn = downloader[di];
                    if(dn === undefined) {
                        print('DATA', di + ',-1,0,0');
                        continue;
                    }
                    var pc;
                    if(totals[di] <= 0) pc = -1;
                    else pc = (dn / totals[di]) * 100;
                    psum += pc;
                    dsum += dn;
                    if(ready.includes(di)) print('DATA', di + ',' + (total == 0 || Math.floor(pc) > 100.0 ? '-1' : Math.floor(pc)) + ',' + totals[di] + ',' + dn);
                    else print('DATA', di + ',-1,0,0');
                }
                print('TOTAL', (!total ? '-1' : total) + ',' + dsum + ',' + (total == 0 || psum < 0 ? '-1' : Math.floor((psum / (100 * trd)) * 100)));
                if(comp >= trd) {
                    clearInterval(statusReporter);
                    if(trd > 1) {
                        print('STATUS', 'MERGING');
                        var s = 'COPY /B ';
                        for(i = 1; i <= trd; i++) s += '"' + fn + '.part.' + i + '"+';
                        s = s.replace(/[+]$/, '');
                        s += ' "' + fn + '"';
                        require('child_process').exec(s, () => {
                            if(Number(process.argv[5]) == 0)
                                for(i = 1; i <= trd; i++) fs.unlinkSync(fn + '.part.' + i, () => 1);
                            print('STATUS', 'COMPLETE');
                            process.exit(0);
                        })
                    } else {
                        fs.renameSync(fn + '.tmp', fn);
                        print('STATUS', 'COMPLETE');
                        process.exit(0);
                    }
                }
            } catch (e) {}
        }, 100);
    })
});
setInterval(() => 1, 987654321);

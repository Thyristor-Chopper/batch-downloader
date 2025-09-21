process.title = '다운로드 부스터';

(function() {
	const readline = require('readline');
	
	const clear = () => process.stdout.write('\033c');
	var _bgcolor = '', _fgcolor = '', _flags = '';
	const write = prpt => process.stdout.write(_flags + _fgcolor + _bgcolor + prpt);
	const print = prpt => write(prpt + '\r\n');
	
	//   0     1     2     3     4     5     6     7
	// 검정, 파랑, 초록, 옥색, 빨강, 자주, 노랑, 하양
	const col = [ 0, 4, 2, 6, 1, 5, 3, 7, 7 ];

	function setcolor(f, b) {
		_flags = f >= 9 ? '\x1b[1m' : '\x1b[0m';
		_bgcolor = '\x1b[4' + (col[b] || 0) + 'm';
		_fgcolor = '\x1b[3' + (col[(f > 9) ? (f - 8) : f] || 0) + 'm';
	}
	
	function gotoxy(x, y) {
		readline.cursorTo(process.stdout, -999, -999);
		readline.cursorTo(process.stdout, x, y);
	}
	
	function moveCursor(x, y) {
		return readline.cursorTo(process.stdout, x, y);
	}
	
	global.setcolor = setcolor;
	global.print = print;
	global.write = write;
	global.gotoxy = gotoxy;
	global.__defineGetter__('cls', () => clrscr());
	global.clrscr = clear;
	global.moveCursor = moveCursor;
})();

const readline = require('readline');
const fs = require('fs');
const child_process = require('child_process');

const _nodeVer = process.version.match(/^v(\d+)[.](\d+)/);
const nodeVer = Number(_nodeVer[1]) + (_nodeVer[2] * 0.1);
if(Number(_nodeVer[1]) != 4) {
	process.exit(0 * console.log('Only Node.js v4 is supported.'));
} else {
	global.async = require('asyncawait').async;
	global.await = require('asyncawait').await;
}

const URL = require('url');
const print = console.log;

if(!console.clear) console.clear = (function() {
	process.stdout.write('\033c');
});

if(!Array.prototype.includes) {
    Array.prototype.includes = (function(fnd) {
        for(var item of this) {
            if(item == fnd) return 1;
        }
        return 0;
    });
}

function input(prpt) {
	return new Promise((resolve, reject) => {
		process.stdout.write(prpt);
		const rl = readline.createInterface(process.stdin, process.stdout);
		rl.question(prpt, ret => {
			rl.close();
			resolve(ret);
		});
	});
}

function timeout(ms) {
	return new Promise((resolve, reject) => {
		setTimeout(() => resolve(0), ms);
	});
}

function progress(val, szz) {
	if(szz == 2) {
		if(val > 95) return '[###############################]';
		if(val > 90) return '[#############################--]';
		if(val > 85) return '[###########################----]';
		if(val > 80) return '[#########################------]';
		if(val > 75) return '[#######################--------]';
		if(val > 70) return '[######################---------]';
		if(val > 65) return '[#####################----------]';
		if(val > 60) return '[###################------------]';
		if(val > 55) return '[#################--------------]';
		if(val > 50) return '[################---------------]';
		if(val > 45) return '[###############----------------]';
		if(val > 40) return '[##############-----------------]';
		if(val > 35) return '[#############------------------]';
		if(val > 30) return '[###########--------------------]';
		if(val > 25) return '[##########---------------------]';
		if(val > 20) return '[#########----------------------]';
		if(val > 15) return '[#######------------------------]';
		if(val > 10) return '[#####--------------------------]';
		if(val >  5) return '[###----------------------------]';
		if(val >  0) return '[#------------------------------]';
		if(val > -1) return '[-------------------------------]';
	} else {
		if(val > 95) return '[####################]';
		if(val > 90) return '[###################-]';
		if(val > 85) return '[##################--]';
		if(val > 80) return '[#################---]';
		if(val > 75) return '[################----]';
		if(val > 70) return '[###############-----]';
		if(val > 65) return '[##############------]';
		if(val > 60) return '[#############-------]';
		if(val > 55) return '[############--------]';
		if(val > 50) return '[###########---------]';
		if(val > 45) return '[##########----------]';
		if(val > 40) return '[#########-----------]';
		if(val > 35) return '[########------------]';
		if(val > 30) return '[#######-------------]';
		if(val > 25) return '[######--------------]';
		if(val > 20) return '[#####---------------]';
		if(val > 15) return '[####----------------]';
		if(val > 10) return '[###-----------------]';
		if(val >  5) return '[##------------------]';
		if(val >  0) return '[#-------------------]';
		if(val > -1) return '[--------------------]';
	}
}

if((process.argv[2] || '').match(/^[%]\d$/)) process.argv[2] = '';
if((process.argv[3] || '').match(/^[%]\d$/)) process.argv[3] = '';
if((process.argv[4] || '').match(/^[%]\d$/)) process.argv[4] = '';

(async(() => {
	var url = process.argv[2] || await (input('파일 주소: '));
	var fn  = process.argv[3] || await (input('파일 이름: '));
	var trd = Number(process.argv[4] || await (input('다운로드 강도: '))) || 1;
	
	fn = fn.replace(/^["]/, '').replace(/["]$/, '');
	
	if(fs.existsSync(fn)) {
		return process.title = '다운로드 부스터 - 수박 씨가 남아...', print('저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오.');
	}
	
	for(i=1; i<=trd; i++) {
		if(fs.existsSync(fn + '.part.' + i)) {
			return process.title = '다운로드 부스터 - 수박 씨가 남아...', print('내부 작업을 위한 파일명이 사용 중입니다. 다른 이름을 선택하십시오.');
		}
	}
	
	const http = require(url.startsWith('https:') ? 'https' : 'http');
	print('리다이렉트 확인 중...');
	
	http.get(url.replace(/^["]/, '').replace(/["]$/, ''), async (res => {
		if(res.headers['location'])
			url = res.headers['location'], print('실제 파일 주소: ' + url);
		
		print('파일을 검사합니다...');
		
		http.get(url, async (res => {
			res.setEncoding('base64'); 
			const total = Number(res.headers['content-length']);
			const boostable = res.headers['accept-ranges'] == 'bytes';
			
			if(!boostable) return process.title = '다운로드 부스터 - 루나 탐사선을 타고...', print('파일 서버가 다운로드를 부스트를 지원하지 않습니다.\n');
			if(!total) return process.title = '다운로드 부스터 - 루나 탐사선을 타고...', print('파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다.\n');
			
			var completed = [], comp = 0;
			var downloader = [];
			var downloads = [];
			var totals = [];
			var tt = [];
			// trd = (total % trd ? (trd + 1) : trd);
			var unit = Math.floor(total / trd);
			var range = 0;
			
			function get(opt) {
				return new Promise((resolve, reject) => {
					http.get({
						host: URL.parse(url).host,
						path: URL.parse(url).path,
						headers: {
							'Range': 'bytes=' + range + '-' + (range + unit)
						}
					}, res => resolve(res));
				});
			}
			
			var ready = [];
			
			print('다운로드를 시작합니다...');
			
			(async(function() {
				for (
					i = 1, range = 0;
					i <= trd; 
					i++  // , range += (unit + 0)
				) {
					(function() {
						var res = await (get({
							host: URL.parse(url).host,
							path: URL.parse(url).path,
							headers: {
								'Range': 'bytes=' + range + '-' + (range + unit)
							}
						}));
						// res.setEncoding('binary');
						
						if((res.statusCode + '')[0] != 2) return;
						
						var id = i;
						
						ready.push(i);
						// print('다운로드 #' + i + ' 준비.');
						
						downloader[id] = 0;
						downloads[id] = '';
						completed[id] = 0;
						totals[id] = tt[id] = Number(res.headers['content-length']);
						// print(range, totals[id], range + totals[id])
						range += totals[id];
						
						res.on('error', () => 1);
						res.on('data', chunk => (downloader[id] += chunk.length, fs.appendFileSync(fn + '.part.' + id, chunk)));
						res.on('end', () => comp++, completed[id] = 1);
						
						// set(i);
					})();
					
					await (timeout(100));
				}
			}))();
			
			console.clear();
			
			var printer = setInterval(async(() => {
				try {
				var totalbytes = '';
				var prt = '';
				var psum = 0;
				var dsum = 0;
				for(di=1; di<=trd; di++) {
					if(di == 'includes') continue;
					var dn = downloader[di];
					if(dn === undefined) {
						prt += ('다운로더 ' + (di < 10 ? ' ' : '') + '#' + di + ': 시작 중...\n');
						continue;
					}
					var pc = (dn / totals[di]) * 100;
					psum += pc;
					dsum += dn;
					
					if(ready.includes(di))
						prt += ('다운로더 ' + (di < 10 ? ' ' : '') + '#' + di + ': ' + progress(pc) + ' (' + Math.floor(pc) + '%) ' + totals[di] + ' 중 ' + dn + ' 완료\n');
					else
						prt += ('다운로더 ' + (di < 10 ? ' ' : '') + '#' + di + ': 시작 중...\n');
				} gotoxy(0, 0); 
				process.title = '다운로드 부스터 - 푸른 나무에서 장작불로 (' + Math.floor((psum / (100 * trd)) * 100) + '%)';
				prt = '총 ' + total + ' 중 ' + dsum + ' 완료 ' + progress(Math.floor((psum / (100 * trd)) * 100), 2) + ' (' + Math.floor((psum / (100 * trd)) * 100) + '%)\n\n' + prt;
				print(prt);
				
				if(comp >= trd) {
					clearInterval(printer);
					
					print('\n파일 조각 결합 중...');
					var s = 'COPY /B ';
					for(i=1; i<=trd; i++) {
						print('  ' + i + '/' + trd + '번 파일 추가');
						s += '"' + fn + '.part.' + i + '"+';
					}
					s = s.replace(/[+]$/, '');
					s += ' "' + fn + '"';
					child_process.exec(s, () => {
						//print('\n임시 파일 삭제 증...');
						for(i=1; i<=trd; i++) {
							//print('  ' + i + '/' + trd + '번 파일 삭제');
							//fs.unlinkSync(fn + '.part.' + i, () => 1);
						}
						print('\n다운로드 끝!');
						process.exit(0);
					});
				}
				} catch(e) {
					print(e.stack);
				}
			}), 100);
		})).end();
	}));
}))();

setInterval(() => 5, 20202020);

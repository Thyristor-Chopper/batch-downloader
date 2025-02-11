Attribute VB_Name = "booster_js"
Private Const CRYPT_STRING_BASE64 As Long = 1
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long

Public CachePath As String

Const BoosterJS As String = _
    "const fs=require('fs');const URL=require('url');if(!Array.prototype.includes)Array.prototype.includes=function includes(find){for(var item of this)if(item==find)return!0;return!1};function timeout(ms){return new Promise((resolve,reject)=>{setTimeout(()=>resolve(0),ms)})}function print(){return console.log.apply(this,Array.prototype.slice.call(arguments).concat(['\r']))}if((process.argv[2]||'').match(/^[%]\d$/))process.argv[2]='';if((process.argv[3]||'').match(/^[%]\d$/))process.argv[3]='';if((process.argv[4]||'').match(/^[%]\d$/))process.argv[4]='';var url=process.argv[2]||process.exit(2);var fn=process.argv[3]||process.exit(2);var trd=Number(process.argv[4]||process.exit(3))||process.exit(3);fn=fn.replace(/^[""]/,'').replace(/[""]$/,'');const parsed=require('path').parse(fn);if(fs.existsSync(fn)){if(Number(process.argv[6])==0)return process.exit(4);else if(Number(process.argv[6])==1)fs.unlinkSync(fn,()=>1);else if(Number(process.argv[6])==2)fn=parsed.dir.replace(/\\$/,'')+'\\'+parsed.name+'-'+" & _
    "Math.floor(Math.random()*10000000000000000)+parsed.ext;print('MODIFIEDFILENAME',fn)}if(fn.endsWith('.')){fn=fn.replace(/[.]$/,'_');print('MODIFIEDFILENAME',fn)}if(trd<=1&&fs.existsSync(fn+'.tmp'))fs.unlinkSync(fn+'.tmp');for(var i=1;i<=trd;i++)if(fs.existsSync(fn+'.part.'+i))return process.exit(5);const http=require(url.startsWith('https:')?'https':'http');print('STATUS','CHECKREDIRECT');http.get(url.replace(/^[""]/,'').replace(/[""]$/,''),res=>{if(res.headers.location)url=res.headers.location,print('REALADDR',url);if(trd>1)print('STATUS','CHECKFILE');http.get(url,res=>{res.setEncoding('base64');var total=Number(res.headers['content-length']);if(trd>1){if(res.headers['accept-ranges']!='bytes')return process.exit(6);if(!total)return process.exit(7);}else if(!total){total=0}var completed=[],comp=0;var downloader=[];var downloads=[];var totals=[];var tt=[];var unit=Math.floor(total/trd);var range=0;function get(){var headers={};if(trd>1)headers.Range='bytes='+range+'-'+(range+unit);" & _
    "return new Promise((resolve,reject)=>{http.get({host:URL.parse(url).host,path:URL.parse(url).path,headers,},res=>resolve(res))})}var ready=[];print('STATUS','DOWNLOADING');(async function(){for(var i=1,range=0;i<=trd;i++){const response=await get();if((response.statusCode+'')[0]!=2){await timeout(100);continue}const id=i;ready.push(i);downloader[id]=0;downloads[id]='';completed[id]=0;totals[id]=tt[id]=Number(response.headers['content-length']||0);range+=totals[id];response.on('error',()=>1);response.on('data',chunk=>(downloader[id]+=chunk.length,fs.appendFileSync(fn+(trd<=1?'.tmp':('.part.'+id)),chunk)));response.on('end',()=>comp++,completed[id]=1);await timeout(100)}})();var statusReporter=setInterval(async()=>{try{var totalbytes='';var prt='';var psum=0;var dsum=0;for(di=1;di<=trd;di++){if(di=='includes')continue;var dn=downloader[di];if(dn===undefined){print('DATA',di+',-1,0,0');continue}var pc;if(totals[di]<=0)pc=-1;else pc=(dn/totals[di])*100;psum+=pc;dsum+=dn;" & _
    "if(ready.includes(di))print('DATA',di+','+(total==0||Math.floor(pc)>100.0?'-1':Math.floor(pc))+','+totals[di]+','+dn);else print('DATA',di+',-1,0,0')}print('TOTAL',(!total?'-1':total)+','+dsum+','+(total==0||psum<0?'-1':Math.floor((psum/(100*trd))*100)));if(comp>=trd){clearInterval(statusReporter);if(trd>1){print('STATUS','MERGING');var s='COPY /B ';for(i=1;i<=trd;i++)s+='""'+fn+'.part.'+i+'""+';s=s.replace(/[+]$/,'');s+=' ""'+fn+'""';require('child_process').exec(s,()=>{if(Number(process.argv[5])==0)for(i=1;i<=trd;i++)fs.unlinkSync(fn+'.part.'+i,()=>1);print('STATUS','COMPLETE');process.exit(0)})}else{fs.renameSync(fn+'.tmp',fn);" & _
    "print('STATUS','COMPLETE');process.exit(0)}}}catch(e){}},100)})});setInterval(()=>1,987654321)"

Dim NodeBase64(25332) As String

'https://www.vbforums.com/showthread.php?879111-vb-6-0-convert-base64-image-data-into-bmp-image
Private Function atob(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    If lSize > 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        atob = baOutput
    Else
        atob = vbNullString
    End If
End Function

'https://stackoverflow.com/questions/10725102/how-to-correctly-write-the-contents-of-a-byte-array-to-a-file-in-vb6
Private Sub WriteFromByte(Bytes() As Byte, FileName As String)
    Dim fnum As Integer
    fnum = FreeFile()
    Open FileName For Binary As #fnum
    Put #fnum, 1, Bytes
    Close fnum
End Sub

Private Sub WriteFromString(str As String, FileName As String)
    Dim fnum As Integer
    fnum = FreeFile()
    Open FileName For Binary As #fnum
    Put #fnum, 1, str
    Close fnum
End Sub

Private Sub AppendFromByte(Bytes() As Byte, FileName As String)
    Dim fnum As Integer
    fnum = FreeFile()
    Open FileName For Binary Access Write As #fnum
    Seek #fnum, LOF(fnum) + 1
    Put #fnum, , Bytes
    Close #fnum
End Sub

Sub LoadJS()
    On Error Resume Next
    MkDir CachePath
    On Error GoTo 0
    If Not FileExists(CachePath & "booster.js") Then
        WriteFromString BoosterJS, CachePath & "booster.js"
    End If
    If Not FileExists(CachePath & "node.exe") Then
		NodeBase64(0) = "TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAEAAA4fug4AtAnNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAAC2JfRr8kSaOPJEmjjyRJo4RthrONxEmjhG2Gk4MESaOEbYaDjBRJo4bORdOPhEmjjJGpk56USaOMkanzmzRJo4yRqeOdZEmjjyRJo480SaOGAamTngRJo4L7tROOFEmjjyRJs40UWaOGUakzknQZo4ZRqaOfNEmjhgGmU480SaOPJEDTjxRJo4ZRqYOfNEmjhSaWNo8kSaOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFBFAABMAQcATK4IYQAAAAAAAAAA4AACAQsBDgAAoJgAAF6MAAAAAACbrZQAABAAAACwmAAAAEAAABAAAAACAAAFAAEAAAAAAAUAAQAAAAAAAEAlAQAEAAAAAAAAAwBAgQAAEAAAEAAAAAAQAAAQAAAAAAAAEAAAAGAaCgGUOBAA9FIaAcgAAAAA8BwBCK8BAAAAAAAAAAAAAAAAAAAAAAAAoB4BWJYGACADCQFUAAAAAAAAAAAAAAAAAAAAAAAAANQDCQEYAAAAeAMJAUAAAAAAAAAAAAAAAACwmACMBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALnRleHQAAAAEn5gAABAAAACgmAAABAAAAAAAAAAAAAAAAAAAIAAAYC5yZGF0YQAACLyBAACwmAAAvoEAAKSYAAAAAAAAAAAAAAAAAEAAAEAuZGF0YQAAAFBQAgAAcBoBADYBAABiGgEAAAAAAAAAAAAAAABAAADALmdmaWRzAADkAgAAANAcAQAEAAAAmBsBAAAAAAAAAAAAAAAAQAAAQC50bHMAAAAADQAAAADgHAEAAgAAAJwbAQAAAAAAAAAAAAAAAEAAAMAucnN"
		NodeBase64(1) = "yYwAAAAivAQAA8BwBALABAACeGwEAAAAAAAAAAAAAAABAAABALnJlbG9jAABYlgYAAKAeAQCYBgAATh0BAAAAAAAAAAAAAAAAQAAAQgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFVTVlcz0pxYi8g1AAAgAFCdnFgzyDPAi3QkFIlGCA+64RUPg2sBAAAPoov4M8CB+0dlbnUPlcCL6IH6aW5lSQ+VwAvogfludGVsD5XAC+h0fYH7QXV0aA+VwIvwgfplbnRpD5XAC/CB+WNBTUQPlcAL8HVauAAAAIAPoj0BAACAckyL8LgBAACAD6IL6YHlAQgAAIH+CAAAgHIzuAgAAIAPog+28Ua4AQAAADPJD6IPuuIcD4ODAAAAwesQgeP/AAAAO953doHi////7+tug/8Evv////9yF7gEAAAAuQAAAAAPoovwwe4Ogeb/DwAAuAEAAAAzyQ+igeL//++/g/0AdRSBygAAAECA5A+A/A91BoHKAAAQAA+64hxzH4Hi////74P+AHQUgcoAAAAQwesQgPsBdwaB4v///++B5QAIAACB4f/3//+L8gvpg/8Hi3wkFHIMuAcAAAAzyQ+iiV8ID7rlG3MeM8kPAdCD4AaD+AZ0G4P4AnQMgeX9///9geb////+geX/5//vg2cI34vGi9VfXltdw42kJAAAAACQM8Az0o0NgL9cAQ+6IQRzAg8xw42kJAAAAACNmwAAAACNDYC/XAEPuiEEcyMOkFiQg+ADdRqcWA+64AlzEg8xUlD0DzErBC"

        AppendFromByte atob(NodeBase64(0)), CachePath & "node.exe"
        AppendFromByte atob(NodeBase64(1)), CachePath & "node.exe"
    End If
    
    Exit Sub
End Sub


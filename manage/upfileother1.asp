<!--#include file="include/db_conn.asp"-->
<link href="inc/southidc.css" rel="stylesheet" type="text/css">
<%sub error2(message)%>
<script>alert('<%=message%>');history.back();</script><script>window.close();</script>
<%end sub
dim oUpFileStream

Class Upload_file

dim Form,File,Err

Private Sub Class_Initialize
Err=-1
end sub

Private Sub Class_Terminate 
'清除变量及对像
if Err < 0 then
oUpFileStream.Close
Form.RemoveAll
File.RemoveAll
set Form=nothing
set File=nothing
set oUpFileStream =nothing
end if
End Sub

Public Sub GetDate(RetSize)
'定义变量
dim RequestBinDate,sStart,bCrLf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,oFileInfo
dim iFileSize,sFilePath,sFileType,sFormvalue,sFileName
dim iFindStart,iFindEnd
dim iFormStart,iFormEnd,sFormName
'代码开始
If Request.TotalBytes < 1 Then
Err=1
Exit Sub
End If
If RetSize > 0 Then 
If Request.TotalBytes > RetSize then
Err=2
Exit Sub
End If
End If
set Form = Server.CreateObject("Scripting.Dictionary")
set File = Server.CreateObject("Scripting.Dictionary")
set tStream = Server.CreateObject("adodb.stream")
set oUpFileStream = Server.CreateObject("adodb.stream")
oUpFileStream.Type = 1
oUpFileStream.Mode = 3
oUpFileStream.Open 
oUpFileStream.Write Request.BinaryRead(Request.TotalBytes)
oUpFileStream.Position=0
RequestBinDate = oUpFileStream.Read 
iFormEnd = oUpFileStream.Size
bCrLf = chrB(13) & chrB(10)
'取得每个项目之间的分隔符
sStart = MidB(RequestBinDate,1, InStrB(1,RequestBinDate,bCrLf)-1)
iStart = LenB (sStart)
iFormStart = iStart+2
'分解项目
Do
iInfoEnd = InStrB(iFormStart,RequestBinDate,bCrLf & bCrLf)+3
tStream.Type = 1
tStream.Mode = 3
tStream.Open
oUpFileStream.Position = iFormStart
oUpFileStream.CopyTo tStream,iInfoEnd-iFormStart
tStream.Position = 0
tStream.Type = 2
tStream.Charset ="UTF-8"
sInfo = tStream.ReadText 
'取得表单项目名称
iFormStart = InStrB(iInfoEnd,RequestBinDate,sStart)-1
iFindStart = InStr(22,sInfo,"name=""",1)+6
iFindEnd = InStr(iFindStart,sInfo,"""",1)
sFormName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
'如果是文件
if InStr (45,sInfo,"filename=""",1) > 0 then
set oFileInfo= new FileInfo
'取得文件属性
iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
iFindEnd = InStr(iFindStart,sInfo,"""",1)
sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
oFileInfo.FileName = GetFileName(sFileName)
oFileInfo.FilePath = GetFilePath(sFileName)
oFileInfo.FileExt = GetFileExt(sFileName)
iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
iFindEnd = InStr(iFindStart,sInfo,vbCr)
oFileInfo.FileType = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
oFileInfo.FileStart = iInfoEnd
oFileInfo.FileSize = iFormStart -iInfoEnd -2
oFileInfo.FormName = sFormName
file.add sFormName,oFileInfo
else
'如果是表单项目
tStream.Close
tStream.Type = 1
tStream.Mode = 3
tStream.Open
oUpFileStream.Position = iInfoEnd 
oUpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
tStream.Position = 0
tStream.Type = 2
tStream.Charset = "UTF-8"
sFormvalue = tStream.ReadText 
form.Add sFormName,sFormvalue
end if
tStream.Close
iFormStart = iFormStart+iStart+2
'如果到文件尾了就退出
loop until (iFormStart+2) = iFormEnd 
RequestBinDate=""
set tStream = nothing
End Sub

'取得文件路径
Private function GetFilePath(FullPath)
If FullPath <> "" Then
GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
Else
GetFilePath = "../pic/"
End If
End function

'取得文件名
Private function GetFileName(FullPath)
If FullPath <> "" Then
GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
Else
GetFileName = ""
End If
End function

'取得扩展名
Private function GetFileExt(FullPath)
If FullPath <> "" Then
GetFileExt = mid(FullPath,InStrRev(FullPath, ".")+1)
Else
GetFileExt = ""
End If
End function

End Class

'文件属性类
Class FileInfo
dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
Private Sub Class_Initialize 
FileName = ""
FilePath = ""
FileSize = 0
FileStart= 0
FormName = ""
FileType = ""
FileExt = ""
End Sub

'保存文件方法
Public function SaveToFile(FullPath)
dim oFileStream,ErrorChar,i
SaveToFile=1
if trim(fullpath)="" or right(fullpath,1)="/" then exit function
set oFileStream=CreateObject("Adodb.Stream")
oFileStream.Type=1
oFileStream.Mode=3
oFileStream.Open
oUpFileStream.position=FileStart
oUpFileStream.copyto oFileStream,FileSize
oFileStream.SaveToFile FullPath,2
oFileStream.Close
set oFileStream=nothing 
SaveToFile=0
end function

'取得文件内容
Public Function GetDate
oUpFileStream.Position =FileStart
GetDate=oUpFileStream.Read(FileSize)
End Function
End Class
%>
<%
if Request("menu")="up" then
On Error Resume Next
if request("atype")<>"" then
atype=request("atype")
end if

set FileUP=new Upload_file 
FileUP.GetDate(-1)
formPath="../pic/"
set file=FileUP.file("file")
filename=formPath&right(year(now),2)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"."&file.FileExt
if file.filesize > 153600000 then
error2("文件大小不得超过 13M \n当前的文件大小为 "&int(file.filesize/1024)&" K")
end if
if  LCase(file.FileExt)="gif" or LCase(file.FileExt)="jpg" or LCase(file.FileExt)="swf" or LCase(file.FileExt)="bmp" or LCase(file.FileExt)="doc" or LCase(file.FileExt)="rar" or LCase(file.FileExt)="flv" then 
img=""&filename&""
else
error2("对不起，本服务器只支持GIF、JPG、swf、doc、rar格式的文件\n不支持 "&file.FileExt&" 格式的文件")
response.end
end if
file.SaveToFile Server.mappath(filename)
set FileUP=nothing

%>
<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0>
<link href=inc/css.css rel=stylesheet>
<%if atype<>"" then%><SCRIPT>parent.myform.<%=atype%>.value='<%=replace(filename,"../","")%>'</SCRIPT><%end if%>
<script language=JavaScript>  
function copyCode(o)
{o.select();
var js=o.createTextRange();
js.execCommand("Copy");
}
<%if atype<>"" then%>
document.write("<font color=red>上传成功!</font> 图片链接[双击复制]:");
document.write("<input onfocus='this.select();copyCode(this)' style='width:200;' name='link_img' id='link_img'  value=<%=replace(filename,"../","")%>>");
<%else%>
document.write("<p style='line-height: 150%; margin: 5px'><font color=red>上传成功!</font> 图片链接[双击复制并关闭窗口]:<br>");
document.write("<input onfocus='this.select();' ondblclick='copyCode(this);javascript:window.close()' style='width:200;' value=<%=filename%>>");
<%end if%>
</script><a target=_blank href=<%=filename%>>打开</a> <a href=# onClick=history.go(-1)><font color=#ff0000>重新上传</font></a>
<%response.end
a=filename
else%>
<body topmargin="0" class="a1" leftmargin="0" rightmargin="0" bottommargin="0">
<link href=inc/css.css rel=stylesheet>
<form enctype=multipart/form-data method=post action=upfileother1.asp?menu=up&atype=<%=request("atype")%>>
<table cellpadding=0 cellspacing=0 width=100%>
<tr><td bgcolor="ecf4ff" width="100%"><input type=file style=FONT-SIZE:9pt name="file" size="30"> <input style=FONT-SIZE:9pt type="submit" value=" 上 传 " name=Submit></td></tr></table>
<%end if%>

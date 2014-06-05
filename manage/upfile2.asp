<!--#include file="upload_5xsoft.inc"-->
<%
session("name")=request("name")
Function CheckFileExt(FileExt)
	Dim ForumUpload,i
'##################################################################
'								文件类型
'##################################################################
	ForumUpload="gif,jpg,bmp,jpeg,png,swf,doc"'上传的文件类型
	ForumUpload=Split(ForumUpload,",")'转换成数组
	For i=0 to UBound(ForumUpload)
		If LCase(FileExt)=Lcase(Trim(ForumUpload(i))) Then
			CheckFileExt=True
			Exit Function
		End If
	Next
	if not CheckFileExt then CheckFileExt=False
End Function

Function CheckFileType(FileType)
	CheckFileType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileType = True
	If Cstr(Lcase(Trim(FileType)))="application/x-shockwave-flash" Then CheckFileType = True
	If Cstr(Lcase(Trim(FileType)))="application/octet-stream" Then CheckFileType = True
	If Cstr(Lcase(Trim(FileType)))="application/msword" Then CheckFileType = True
End Function
'##################################################################
'								上传之后的文件名
'##################################################################
function MakedownName()
dim fname
fname = now()
fname = replace(fname,"-","")
fname = replace(fname," ","") 
fname = replace(fname,":","")
fname = replace(fname,"PM","")
fname = replace(fname,"AM","")
fname = replace(fname,"上午","")
fname = replace(fname,"下午","")
fname = int(fname) + int((10-1+1)*Rnd + 1)
MakedownName=fname
end function
'##################################################################
'								文件类型的检查
'##################################################################
Function FixName(UpFileExt)
	If IsEmpty(UpFileExt) Then Exit Function
	FixName = Lcase(UpFileExt)
	FixName = Replace(FixName,Chr(0),"")
	FixName = Replace(FixName,".","")
	FixName = Replace(FixName,"asp","")
	FixName = Replace(FixName,"asa","")
	FixName = Replace(FixName,"aspx","")
	FixName = Replace(FixName,"cer","")
	FixName = Replace(FixName,"cdx","")
	FixName = Replace(FixName,"htr","")
End Function
%>
<%

 Dim objFSO    '声明一个名称为 objFSO 的变量以存放对象实例
 '##################################################################
 '								存放路径
 '##################################################################
formPath="product/"
names=session("name")
set upload=new upload_5xSoft ''建立上传对象

 '##################################################################
 '						检测是否在formPath目录下
 '							在目录后加(/)
 '##################################################################
 if right(formPath,1)<>"/" then formPath=formPath&"/" 
'iCount=0
set file=upload.file("file1")
 '##################################################################
 '						检测是否有上传文件
 '##################################################################
if file.FileSize>0 then
 '##################################################################
 '						检测文件的大小
 '##################################################################
if file.filesize>200000000 then
response.write"<SCRIPT language=JavaScript>alert('您上传的图片大于规定大小(200K),请改变文件大小后再进行上传。');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
 '##################################################################
 '						得到上传文件类型
 '##################################################################
FileExt	= Mid(file.Filename, InStrRev(file.Filename, ".")+1)
 '##################################################################
 '						替换掉非法文件类型
 '##################################################################
FileExt	= FixName(FileExt)
 '##################################################################
 '						检测文件的类型
 '##################################################################
If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
response.write"<SCRIPT language=JavaScript>alert('您上传的文件必须是gif|jpg|jpeg|bmp|png|doc图象文件或压缩文件，请将你上传的文件转换为以上格式后再进行上传。');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
 '##################################################################
 '						为文件命名
 '##################################################################
FileName=MakedownName()&"."&FileExt
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")'注册FSO文件系统组件
  If objFSO.FolderExists(Server.MapPath(""&formPath&"")) Then'如果存在文件夹就直接保存图片
		file.SaveAs Server.mappath(formPath&FileName)
  Else
   objFSO.CreateFolder(Server.MapPath(""&formPath&""))'不存在就建一个目录并保存
		file.SaveAs Server.mappath(formPath&FileName)
  End If
  Set objFSO = Nothing      '释放 FileSystemObject 对象实例内存空间
	response.write"<script>parent.formm."&session("name")&".value='"&FileName&"'</script>"
	Response.Write "<script>alert('上传成功');</script>"
	end if
	set upload=nothing
	session("name")=""
 %>      



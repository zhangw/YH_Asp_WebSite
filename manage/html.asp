<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>中国高力汽配网</title>
<!--#include file="include/db_conn.asp"-->
<!--#include file="../inc/str.asp"-->
<!--#include file="test.asp"-->
<%
Response.Expires = 0
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<link href="css/main.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.STYLE1 {font-size: 12px}
-->
</style>
</head>

<body>
<table cellSpacing="0" cellPadding="0" width="100%" bgColor="#f1f1f1" border="0">
				<tr>
					<td>						<table cellSpacing="0" cellPadding="0" background="images/sheet_bk.gif" border="0" style="margin-top:5px;">
							<tr>
								<td vAlign="bottom"><IMG height="18" src="images/sheet_left.gif" width="24"></td>
								<td class="SheetSelected" vAlign="bottom" width="60">后台管理</td>
								<td valign="bottom"><img src="images/sheet_right.gif" width="25" height="18"></td>
							</tr>
						</table>
					</td>
					<td vAlign="bottom" align="right">
						<table cellSpacing="2" cellPadding="0" width="99%" border="0">
							<tr>
								<td align="right">您的位置：<A href="index.asp" target="_top">后台管理</A> &gt;&gt; 生成页面</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
</table><br>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 生成页面</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<%if request("action")="" then%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="12%" height="25" bgcolor="#e0e9f8"  class="STYLE5">项目</th>
    <th width="12%" bgcolor="#e0e9f8"  class="STYLE5">操作</th>
    <th width="12%" height="25" bgcolor="#e0e9f8"  class="STYLE5">项目</th>
    <th width="12%" bgcolor="#e0e9f8"  class="STYLE5">操作</th>
    <th width="12%" bgcolor="#e0e9f8"  class="STYLE5">项目</th>
    <th width="12%" bgcolor="#e0e9f8"  class="STYLE5">操作</th>
    <th width="12%" bgcolor="#e0e9f8"  class="STYLE5">项目</th>
    <th width="12%" bgcolor="#e0e9f8"  class="STYLE5">操作</th>
  </tr>
  <tr bgcolor="#fffff0"class="unnamed1">
    <td width="13%" height="20" align="center"  class="STYLE5" style="padding-left:5px;">首页</td>

	<td width="13%" align="center"  class="STYLE5"><a href="?action=sy"><img src="images/sc.jpg" width="58" height="19" border="0" /></a></td>
	<td width="13%" align="center" >信息列表页</td>
    <td width="13%" align="center" ><span class="STYLE5"><a href="?action=lb"><img src="images/sc.jpg" width="58" height="19" border="0" /></a></span></td>
    <td width="13%" align="center" >信息页面</td>
    <td width="13%" align="center" ><span class="STYLE5"><img src="images/sc.jpg" width="58" height="19" /></span></td>
    <td width="13%" align="center" >关于乔恩</td>
    <td width="13%" align="center" ><span class="STYLE5"><img src="images/sc.jpg" width="58" height="19" /></span></td>
  </tr>
</table>
<br />
  <br />
  <br />
  <%end if%>
  <%if request("action")="sy" then
IndexTopCode=FSOFileRead("/top.htm")
IndexbootomCode=FSOFileRead("/bootom.htm")
IndexadCode=FSOFileRead("/ad.htm")
IndexCode=FSOFileRead("/mb_index.htm")
IndexCode=Replace(IndexCode,"$content_top$",IndexTopCode)
IndexCode=Replace(IndexCode,"$content_bootom$",IndexbootomCode)
IndexCode=Replace(IndexCode,"$meifabufa$",meifabufa())
IndexCode=Replace(IndexCode,"$content_ad$",IndexadCode)
IndexCode=Replace(IndexCode,"$chanpzhans$",chanpzhans())
IndexCode=Replace(IndexCode,"$hualiaotuofa$",hualiaotuofa())
IndexCode=Replace(IndexCode,"$tuofabufa$",tuofabufa())
IndexCode=Replace(IndexCode,"$bafazhuanti$",bafazhuanti())
IndexCode=Replace(IndexCode,"$gongsigongao$",gongsigongao())

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
Set fout = fso.CreateTextFile(server.mappath("/index.htm")) 
fout.Write IndexCode 
fout.close 
set fout=nothing 
%>
   <div align="center"> <a href='html.asp' class="Sheet STYLE1"><br />
     <br />
     <br />
     <br />
     <br />
   成功生成，点击返回</a></div>
  <%end if
  
 if request("action")="lb" then
IndexTopCode=FSOFileRead("/top.htm")
IndexbootomCode=FSOFileRead("/bootom.htm")
IndexCode=FSOFileRead("/mb_index.htm")
IndexCode=Replace(IndexCode,"$content_top$",IndexTopCode)
IndexCode=Replace(IndexCode,"$content_bootom$",IndexbootomCode)
IndexCode=Replace(IndexCode,"$meifabufa$",meifabufa())
IndexCode=Replace(IndexCode,"$chanpzhans$",chanpzhans())
IndexCode=Replace(IndexCode,"$hualiaotuofa$",hualiaotuofa())
IndexCode=Replace(IndexCode,"$tuofabufa$",tuofabufa())
IndexCode=Replace(IndexCode,"$bafazhuanti$",bafazhuanti())
IndexCode=Replace(IndexCode,"$gongsigongao$",gongsigongao())

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
Set fout = fso.CreateTextFile(server.mappath("/index.htm")) 
fout.Write IndexCode 
fout.close 
set fout=nothing 
%>
   <div align="center"> <a href='html.asp' class="Sheet STYLE1"><br />
     <br />
     <br />
     <br />
     <br />
   成功生成，点击返回</a></div>
  <%end if
   
function FSOFileRead(filename)
Dim objFSO,objCountFile,FiletempData
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objCountFile = objFSO.OpenTextFile(Server.MapPath(filename),1,True)
FSOFileRead = objCountFile.ReadAll
objCountFile.Close
Set objCountFile=Nothing
Set objFSO = Nothing
End Function

function meifabufa()
str="<table width=318 border=0 align=center cellspacing=0><tr><td><table width=100% border=0 cellspacing=0 cellpadding=0 class=Index-Center3>"
set rs = server.CreateObject("adodb.recordset")
sql = "select top 6 * from news where newsid=21 order by id desc"
rs.open sql,conn,1,3
do while not rs.eof
str=str&"<tr ><td width=11></td><td height=22><img src=Images/Arrow-Page-01.gif  align=absmiddle>&nbsp;<a href=NewsView.asp?id="&rs("id")&" target='_blank'>"
if CheckStringLength(rs("title"))>34 then
str=str&InterceptString(rs("title"),34)&"..."
else
str=str&rs("title")
end if
str=str&"</a></td><td width=60 class=Index-Center2>"&formatdatetime(rs("time"),2)&"</td></tr>"
rs.movenext
loop
str=str&"</table></td></tr></table>"
meifabufa=str
End Function

function tuofabufa()
str="<table width=318 border=0 align=center cellspacing=0><tr><td><table width=100% border=0 cellspacing=0 cellpadding=0 class=Index-Center3>"
set rs = server.CreateObject("adodb.recordset")
sql = "select top 6 * from news where newsid=18 order by id desc"
rs.open sql,conn,1,3
do while not rs.eof
str=str&"<tr ><td width=11></td><td height=22><img src=Images/Arrow-Page-01.gif  align=absmiddle>&nbsp;<a href=NewsView.asp?id="&rs("id")&" target='_blank'>"
if CheckStringLength(rs("title"))>34 then
str=str&InterceptString(rs("title"),34)&"..."
else
str=str&rs("title")
end if
str=str&"</a></td><td width=60 class=Index-Center2>"&formatdatetime(rs("time"),2)&"</td></tr>"
rs.movenext
loop
str=str&"</table></td></tr></table>"
tuofabufa=str
End Function

function hualiaotuofa()
str="<table width=318 border=0 align=center cellspacing=0><tr><td><table width=100% border=0 cellspacing=0 cellpadding=0 class=Index-Center3>"
set rs = server.CreateObject("adodb.recordset")
sql = "select top 6 * from news where newsid=23 order by id desc"
rs.open sql,conn,1,3
do while not rs.eof
str=str&"<tr ><td width=11></td><td height=22><img src=Images/Arrow-Page-01.gif  align=absmiddle>&nbsp;<a href=NewsView.asp?id="&rs("id")&" target='_blank'>"
if CheckStringLength(rs("title"))>34 then
str=str&InterceptString(rs("title"),34)&"..."
else
str=str&rs("title")
end if
str=str&"</a></td><td width=60 class=Index-Center2>"&formatdatetime(rs("time"),2)&"</td></tr>"
rs.movenext
loop
str=str&"</table></td></tr></table>"
hualiaotuofa=str
End Function

function chanpzhans()
str="<table width=318 border=0 align=center cellspacing=0><tr><td><table width=100% border=0 cellspacing=0 cellpadding=0 class=Index-Center3>"
set rs = server.CreateObject("adodb.recordset")
sql = "select top 6 * from news where newsid=24 order by id desc"
rs.open sql,conn,1,3
do while not rs.eof
str=str&"<tr ><td width=11></td><td height=22><img src=Images/Arrow-Page-01.gif  align=absmiddle>&nbsp;<a href=NewsView.asp?id="&rs("id")&" target='_blank'>"
if CheckStringLength(rs("title"))>34 then
str=str&InterceptString(rs("title"),34)&"..."
else
str=str&rs("title")
end if
str=str&"</a></td><td width=60 class=Index-Center2>"&formatdatetime(rs("time"),2)&"</td></tr>"
rs.movenext
loop
str=str&"</table></td></tr></table>"
chanpzhans=str
End Function

function bafazhuanti()
str="<table width=318 border=0 align=center cellspacing=0><tr><td><table width=100% border=0 cellspacing=0 cellpadding=0 class=Index-Center3>"
set rs = server.CreateObject("adodb.recordset")
sql = "select top 6 * from news where newsid=25 order by id desc"
rs.open sql,conn,1,3
do while not rs.eof
str=str&"<tr ><td width=11></td><td height=22><img src=Images/Arrow-Page-01.gif  align=absmiddle>&nbsp;<a href=NewsView.asp?id="&rs("id")&" target='_blank'>"
if CheckStringLength(rs("title"))>34 then
str=str&InterceptString(rs("title"),34)&"..."
else
str=str&rs("title")
end if
str=str&"</a></td><td width=60 class=Index-Center2>"&formatdatetime(rs("time"),2)&"</td></tr>"
rs.movenext
loop
str=str&"</table></td></tr></table>"
bafazhuanti=str
End Function

function gongsigongao()
str="<table width=100% border=0 align=center cellspacing=0><tr><td><table width=100% border=0 cellspacing=0 cellpadding=0 class=Index-Center3>"
set rs = server.CreateObject("adodb.recordset")
sql = "select top 6 * from news where newsid=26 order by id desc"
rs.open sql,conn,1,3
do while not rs.eof
str=str&"<tr ><td width=11></td><td height=22><img src=Images/Arrow-Page-01.gif  align=absmiddle>&nbsp;<a href=NewsView.asp?id="&rs("id")&" target='_blank'>"
if CheckStringLength(rs("title"))>34 then
str=str&InterceptString(rs("title"),34)&"..."
else
str=str&rs("title")
end if
str=str&"</a></td><td width=60 class=Index-Center2>"&formatdatetime(rs("time"),2)&"</td></tr>"
rs.movenext
loop
str=str&"</table></td></tr></table>"
gongsigongao=str
End Function
%>
</body>
</html>

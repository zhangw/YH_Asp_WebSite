﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>网站信息管理</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-image: url(images/bss.jpg);
}
-->
</style></head>

<body>
<%
set conn = server.CreateObject("adodb.connection")
conn.ConnectionTimeout = 25
conn.Open "Provider=Microsoft.JET.OLEDB.4.0;" &"Data Source=" & Server.MapPath("../data/db.mdb") & ";Jet OLEDB:Database Password="
if request("action")="login" then
if request("admin") ="" then
response.Write "<script>alert('请输入用户名称...');history.back();</script>"
response.End()
end if
if request("password") ="" then
response.Write "<script>alert('请输入管理密码...');history.back();</script>"
response.End()
end if
if request("code") ="" then
response.Write "<script>alert('请输入验证码...');history.back();</script>"
response.End()
end if
if cint(request("code")) <> cint(Session("GVNUM")) then
response.Write "<script>alert('验证码输入错误...');history.back();</script>"
response.End()
end if
dim rs
admin1=request("admin")
password1=request("password")
set rs=server.CreateObject("ADODB.RecordSet")
rs.open "select * from admin where admin='" & admin1 & "' and password='"& password1 &"'",conn,1,1
if rs.eof and rs.bof then
response.write"<SCRIPT language=JavaScript>alert('用户名或密码不正确！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
else
response.cookies("username")=rs("admin")
response.cookies("password")=rs("password")
response.Redirect("index.asp")
end if
rs.close
set rs=nothing
end if
%>
<table width="645" height="321" border="0" align="center" cellpadding="0" cellspacing="0" background="images/login.jpg" style="margin-top:210px;">
  <tr>
    <td colspan="3" style="font-size:14px; font-family:'黑体'; color:#FFFFFF;"><table cellspacing=0 cellpadding=0 width=400 border=0 style="margin-left:160px;">
      <form action=login.asp?action=login method=post>
        <tbody>
          <tr>
            <td width=93 height="30" align=right>用户名称：</td>
            <td width="134"><font color=#ffffff>
              <input class=input 
                  style="WIDTH: 108px;height:16px;" name=admin TABINDEX=1/>
            </font></td>
            <td width="155" rowspan="4" align="left"><input type="image" src="images/go.jpg" TABINDEX=4></td>
          </tr>
          <tr>
            <td width=93 height="30" align=right>管理密码：</td>
            <td><font color=#ffffff>
              <input class=input 
                  style="WIDTH: 108px;height:16px;" type=password size=18 name=password TABINDEX=2/>
            </font></td>
            <td width="18" rowspan="2">&nbsp;</td
          >
          </tr>
          <tr>
            <td width=93 height="30" align=right>验证码：</td>
            <td valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="43%"><font color=#ffffff>
                    <input name=code type=text class=input id="code" 
                  style="WIDTH: 50px;height:16px;" size=18 TABINDEX=3/>
                  </font></td>
                  <td width="57%"><%
Dim num
Randomize
num = Int(7999 * Rnd + 2000)
Session("GVNUM") = cstr(num)
For i = 1 to Len(num)     
        G_Counts = G_Counts & "<IMG SRC=../gif/" & Mid(num, i, 1) & ".gif>"
Next
Response.Write G_Counts
%></td>
                </tr>
              </table>
              </td>
          </tr>
        </tbody>
      </form>
    </table></td>
  </tr>
</table>
</body>
</html>

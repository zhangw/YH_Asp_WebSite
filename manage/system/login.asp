<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title></title>
<link href="../css/main.css" rel="stylesheet" type="text/css" />
</head>
<%
set conn = server.CreateObject("adodb.connection")
conn.ConnectionTimeout = 25
conn.Open "Provider=Microsoft.JET.OLEDB.4.0;" &"Data Source=" & Server.MapPath("../../data/db.mdb") & ";Jet OLEDB:Database Password="
if request("action")="login" then
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
response.Redirect("../include/start.htm")
end if
rs.close
set rs=nothing
end if
%>
<body>
<table cellSpacing="0" cellPadding="0" width="100%" bgColor="#f1f1f1" border="0">
				<tr>
					<td height="46">&nbsp;</td>
				</tr>
			</table>
			<table cellSpacing="0" cellPadding="0" width="100%" bgColor="#f1f1f1" border="0">
				<tr>
					<td>
						<table cellSpacing="0" cellPadding="0" background="../images/sheet_bk.gif" border="0">
							<tr>
								<td vAlign="bottom"><IMG height="18" src="../images/sheet_left.gif" width="24"></td>
								<td class="SheetSelected" vAlign="bottom" width="60">后台管理</td>
								<td valign="bottom"><img src="../images/sheet_right.gif" width="25" height="18"></td>
							</tr>
						</table>
					</td>
					<td vAlign="bottom" align="right">
						<table cellSpacing="2" cellPadding="0" width="99%" border="0">
							<tr>
								<td align="right">您的位置：<A href="../index.asp" target="_top">后台管理</A> &gt;&gt; 
									管理员登录</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table>
			<table cellSpacing="23" cellPadding="0" width="100%" border="0">
				<tr>
					<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td class="SubTitle" vAlign="bottom"><IMG height="20" src="../images/icon_info.gif" width="20" align="absMiddle">
									管理员登录</td>
							</tr>
							<tr>
								<td>
									<hr width="100%" color="#000000" noShade SIZE="2">
								</td>
							</tr>
							<tr align="center">
							  <td align="center">&nbsp;</td>
							</tr>
						</table>
						<TABLE cellSpacing=3 cellPadding=0 width=260 border=0>
              <FORM action=login.asp?action=login method=post>
              <TBODY>
              <TR>
                <TD align=right width=100>管理员</TD>
                <TD width=3 rowSpan=4></TD>
                <TD><FONT color=#ffffff><INPUT class=input 
                  style="WIDTH: 108px" size=18 name=admin> </FONT></TD></TR>
              <TR>
                <TD align=right width=100>密　码</TD>
                <TD><FONT color=#ffffff><INPUT class=input 
                  style="WIDTH: 108px" type=password size=18 name=password> 
                  </FONT></TD></TR>
                            <TR>
                <TD width=100>&nbsp;</TD>
                <TD vAlign=bottom height=28><INPUT class=button type=submit value=" 登 录 " name=Submit>                </TD></TR></FORM></TABLE>
						<FONT face="宋体">&nbsp;</FONT>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td>&nbsp;</td>
							</tr>
						</table>
						<table cellSpacing="0" cellPadding="4" width="100%" align="center" border="0">
							<tr>
								<td width="29" height="20"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
</body>
</html>


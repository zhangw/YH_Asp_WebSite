<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title></title>
<!--#include file="include/db_conn.asp"-->
<!--#include file="test.asp"-->

<link href="css/main.css" rel="stylesheet" type="text/css">
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
								<td align="right">您的位置：<A href="index.asp" target="_top">后台管理</A> &gt;&gt;管理员管理</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>
<% if Trim(Request.QueryString("action"))="list" then %>
<p>&nbsp;</p>
<table width="90%" border=0 align=center cellpadding="1" cellSpacing=1 bgcolor="#CCCCCC" class="navi">
    <tr align="center"> 
      <th width="100" bgcolor="#e0e9f8"  class="front3">姓名</th>
      <th width="100" height="25" bgcolor="#e0e9f8" class="front3">用户名</th>
    <th width="100" bgcolor="#e0e9f8" class="front3">是/否删除</th>
    </tr>


<%
sql="select * from admin"
set rs=conn.execute(sql)
i=0
do while not rs.eof 
%>
    <tr align="center" <%if i mod 2=0 then%>bgcolor="#fffff0"<%else%>bgcolor="#ffffff"<%end if%>> 
      <td width="100"  class="front3"><%= rs("name") %></td>
      <td width="100"  class="front3"><%= rs("admin") %></td>
      <td width="100"  class="front3"><a href='javascript:if(confirm("确实要删除吗?"))location="?action=del&id=<%=rs("id")%>&lb=<%=request("lb")%>"'>[删除]</a> <a href="?action=member&id=<%=rs("id")%>" class="link1">[修改]</a> </td>
    </tr>
	<% 
  i=i+1
  rs.movenext
  loop
   %>
</table>
  
 <% End If %>

<%
if Trim(Request.QueryString("action"))="add" then
	dim admin,password1,password2,name
	name=Trim(Request.Form("name"))
	admin=Trim(Request.Form("admin"))
	password1=Trim(Request.Form("password1"))
	password2=Trim(Request.Form("password2"))
	
'''''''''''''''''''''''''''''''''''''''''''''''	
if Trim(Request.Form("submit"))="添 加" then
	set rs=server.CreateObject("ADODB.Recordset")
		sql="select * from admin"
		rs.open sql,conn,1,3
		rs.addnew
		rs("admin")=admin
		rs("password")=password2
		rs("name")=name
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		Response.Write "<script>alert('添加成功!');location='?action=list'</script>"
end if
''''''''''''''''''''''''''''''''''''''''''''''
 %>
<script language="JavaScript" type="text/JavaScript">
// 验证用户名和留言
function check_add(){
	var notnull;
	notnull=true;
	if (document.form1.admin.value==""){
		alert("用户名不能为空！");
		document.form1.admin.focus();
		notnull=false;
		}
	else
	if (document.form1.password1.value==""){
		alert("密码不能为空！");
		document.form1.password1.focus();
		notnull=false;
		}
	else
	if (document.form1.password2.value==""){
		alert("确认密码不能为空！");
		document.form1.password2.focus();
		notnull=false;
		}
	else
	if(document.form1.password1.value != document.form1.password2.value){
		alert("两次密码输入不一致！");
		document.form1.password2.focus();
		notnull=false;
		}
	else
	if (document.form1.password2.value.length < 6){
		alert("密码小于6位！");
		document.form1.password2.focus();
		notnull=false;
		}
	else
	if (document.form1.name.value==""){
		alert("姓名不能为空！");
		document.form1.name.focus();
		notnull=false;
		}		
	return notnull;
	}
</script>
<p>&nbsp;</p>
<table width="90%" border=0 align=center cellpadding="1" cellSpacing=1 bgcolor="#CCCCCC" class="navi">
        <form name="form1" method="post" action="?action=add" onSubmit="return check_add();">
          <tr align="center"> 
            <th height="25" colspan="2" bgcolor="#E0E9F8" >添加管理员</th>
          </tr>
          <tr> 
            <td width="80" align="center" bgcolor="#FFFFFF" class="front3">用户名：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="admin" type="text" id="admin" size="15"></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="front3">姓&nbsp;&nbsp;名：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="name" type="text" id="name" size="15"></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#FFFFFF" class="front3">密&nbsp;&nbsp;码：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="password1" type="password" id="password1" size="15"></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="front3">确认密码：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="password2" type="password" id="password2" size="15"></td>
          </tr>

          <tr> 
            <td colspan="2" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="添 加"> 
            &nbsp; <input type="reset" name="Submit2" value="重 置"></td>
          </tr>

        </form>
</table>
<br>
<% End If %>



<%
if Trim(Request.QueryString("action"))="edit" then
if Trim(Request.Form("submit"))="修 改" then
	password=Trim(Request.Form("password"))
	password1=Trim(Request.Form("password1"))
	password2=Trim(Request.Form("password2"))
	set rs=server.CreateObject("ADODB.Recordset")
		sql="select * from admin where admin='" & session("username")&"'"
		rs.open sql,conn,1,3
		rs("password")=password2
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		Response.Write "<script>alert('修改成功!');location='?action=list'</script>"
end if
'''''''''''''''''''''''''''''''''''''
 %>
<script language="JavaScript" type="text/JavaScript">
// 验证用户名和留言
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.password.value==""){
		alert("原密码不能为空！");
		document.form1.password.focus();
		notnull=false;
		}
	else
	if (document.form1.password1.value==""){
		alert("新密码不能为空！");
		document.form1.password1.focus();
		notnull=false;
		}
	else
	if (document.form1.password2.value==""){
		alert("确认密码不能为空！");
		document.form1.password2.focus();
		notnull=false;
		}
	else
	if(document.form1.password1.value != document.form1.password2.value){
		alert("两次密码输入不一致！");
		document.form1.password2.focus();
		notnull=false;
		}
	else
	if (document.form1.password2.value.length < 6){
		alert("密码小于6位！");
		document.form1.password2.focus();
		notnull=false;
		}
	return notnull;
	}
</script>
<p>&nbsp;</p>

<table width="90%" border=0 align=center cellpadding="1" cellSpacing=1 bgcolor="#CCCCCC" class="navi">
        <form name="form1" method="post" action="?action=edit" onSubmit="return check_edit();">
          <tr align="center"> 
            <th height="25" colspan="2" bgcolor="#E0E9F8" >更改密码</th>
          </tr>
          <tr> 
            <td width="90" align="center" bgcolor="#FFFFFF" class="front3">原密码：</td>
            <td  bgcolor="#FFFFFF" class="front3"><input name="password" type="password" id="password" size="15"></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#FFFFFF" class="front3">新密码：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="password1" type="password" id="password1" size="15"></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#FFFFFF" class="front3">确认密码：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="password2" type="password" id="password2" size="15"></td>
          </tr>
          <tr> 
            <th colspan="2" align="left" bgcolor="#FFFFFF" style="padding-left:80px;">
              <input type="submit" name="Submit" value="修 改"> 
          &nbsp; <input type="reset" name="Submit2" value="重 置"></th></tr>
        </form>
</table>
<div align="center">
  <% End If %>


  <% 
if Trim(Request.QueryString("action"))="del" then
	if Trim(Request.Form("submit"))="删 除" then
		id=Trim(Request.QueryString("id"))
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from admin where id="&id
		rs.open sql,conn,1,3
		if request.cookies("username")=rs("admin") then
			Response.Write "<script>alert('不能删除正在登录的账户!');location='?action=list'</script>"
		else
			rs.delete
			rs.update
			rs.requery
			rs.close
			set rs=nothing
			Response.Write "<script>alert('删除成功!');location='?action=list'</script>"
		end if
	end if
 %>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <span class="front3">确认删除<font color="#FF0000">用户</font>吗？</span><br>
</div>
<form name="form1" method="post" action="?action=del&id=<%= Trim(Request.QueryString("id")) %>">
  <TABLE width=200 border=0 align="center" cellPadding=0 
cellSpacing=0 borderColor=#BEBFD9 class=table_out>
    <TR align="center"> 
      <TD height=25 bgcolor="#D7F0FF"><input type="submit" name="Submit" value="删 除"></TD>
      <TD bgcolor="#D7F0FF"><input type="reset" name="Submit2" value="取 消" onClick="javascript:history.go(-1)"></TD>
    </TR>
  </TABLE>
  <div align="center"></div>
</form>
<% End If %>



<%
if Trim(Request.QueryString("action"))="member" then
	name=Trim(Request.Form("name"))
	admin=Trim(Request.Form("admin"))
	password1=Trim(Request.Form("password1"))
	password2=Trim(Request.Form("password2"))
	
'''''''''''''''''''''''''''''''''''''''''''''''	
if Trim(Request.Form("submit"))="修改" then
	set rs=server.CreateObject("ADODB.Recordset")
		sql="select * from admin where id=" & request("id")
		rs.open sql,conn,1,3
		rs("admin")=admin
		if password1<>"" then
		rs("password")=password2
        end if
		rs("name")=name
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		Response.Write "<script>alert('修改成功!');location='?action=list'</script>"
end if
''''''''''''''''''''''''''''''''''''''''''''''
 %>
<script language="JavaScript" type="text/JavaScript">
// 验证用户名和留言
function check_add(){
	var notnull;
	notnull=true;
	if (document.form1.admin.value==""){
		alert("用户名不能为空！");
		document.form1.admin.focus();
		notnull=false;
		}
	else
	if (document.form1.name.value==""){
		alert("姓名不能为空！");
		document.form1.name.focus();
		notnull=false;
		}
	else
	if(document.form1.password1.value != document.form1.password2.value){
		alert("两次密码输入不一致！");
		document.form1.password2.focus();
		notnull=false;
		}
	else
	if (document.form1.password2.value.length < 6){
		alert("密码小于6位！");
		document.form1.password2.focus();
		notnull=false;
		}	
	return notnull;
	}
</script>

<%
	set rs=server.CreateObject("ADODB.Recordset")
		sql="select * from admin where id=" & request("id")
		rs.open sql,conn,1,3

%><p>&nbsp;</p>
<table width="90%" border=0 align=center cellpadding="1" cellSpacing=1 bgcolor="#CCCCCC" class="navi">
        <form name="form1" method="post" action="?action=member&id=<%=request("id")%>" onSubmit="return check_add();">
          <tr align="center"> 
            <th height="25" colspan="2" bgcolor="#E0E9F8" >修改管理员</th>
          </tr>
          <tr> 
            <td width="14%" align="center" bgcolor="#FFFFFF" class="front3">用户名：</td>
            <td width="86%" bgcolor="#FFFFFF" class="front3"><input name="admin" type="text" id="admin" size="15" value="<%=rs("admin")%>"></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="front3">姓&nbsp;&nbsp;名：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="name" type="text" id="name" size="15" value="<%=rs("name")%>"></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#FFFFFF" class="front3">密&nbsp;&nbsp;码：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="password1" type="password" id="password1" size="15"></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="front3">确认密码：</td>
            <td bgcolor="#FFFFFF" class="front3"><input name="password2" type="password" id="password2" size="15"></td>
          </tr>

          <tr> 
            <td colspan="2" align="center" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="修改"> 
            &nbsp; <input type="reset" name="Submit2" value="重 置"></td>
          </tr>

        </form>
</table>
<br>
<% End If %></body>
</html>

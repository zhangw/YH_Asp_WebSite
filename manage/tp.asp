<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="include/db_conn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>编辑出团表</title>
<link href="../css.css" rel="stylesheet" type="text/css" />
<script language=javascript> 
function OnClick(aa){ 
window.returnValue=aa; 
window.close(); 
} 
</script> 
<style type="text/css">
<!--
.style3 {color: #000000}

body,td,th {
	font-size: 12px;
	color: #000000;
}
a {
	font-size: 12px;
	color: #000000;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #000000;
}
a:hover {
	text-decoration: none;
	color: #000000;
}
a:active {
	text-decoration: none;
	color: #000000;
}
-->
</style>
</head>
<base target="_self">
<body>
<br />
<table width="480" height="30" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
  <form id="formm" name="formm" method="post" action="?action=add&pid=<%=request("id")%>">
	
  <tr >
    <td height="20" bgcolor="#FFFFFF" class="front3" ><input name="image" type="hidden" id="image"/>
      &nbsp;
      <iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
    <td width="72" height="20" bgcolor="#FFFFFF" class="front3" ><input type="submit" name="Submit" value="添加" />
      &nbsp;<a href="#"></a></td>
	</tr>
  </form>
</table>
<br />
<table width="480" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr>
    <td height="1"></td>
  </tr>
</table> 
 <%
  		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from tp  order by id asc"
		xx=0
		rs1.open sql,conn,1,3
		if rs1.eof then
		else
			rs1.pagesize=12
	if request("page")<>"" then 	page=cint(Trim(Request.QueryString("page")))
	if page<1 then
		page=1
	elseif page>rs1.pagecount then
		page=rs1.pagecount
	end if
	rs1.AbsolutePage=page
rowCount = rs1.PageSize
	%>
<table width="480" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  <tr > 
<%
do while not rs1.eof and rowcount>0
if i mod 4=0 then response.Write "</tr><tr>"
%>	
 
    <td width="61" height="20" class="front3" > &nbsp;
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
        <form id="form1" name="form1" method="post" action="?action=save&pid=<%=request("id")%>&id=<%=rs1("id")%>"><tr>
          <td><img src="product/<%=rs1("tp")%>" width="120" height="90" /></td>
        </tr>
		 <tr>
          <td height="20" align="center" bgcolor="#FFFFFF"><a href="?action=del&pid=<%=request("id")%>&id=<%=rs1("id")%>">删除</a></td>
        </tr></form>
      </table></td>
 <%
 i=i+1
 RowCount=RowCount-1
		rs1.movenext
  loop
 %> </tr>
</table>
<br />
<table width="480" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr>
    <td height="1"></td>
  </tr>
</table>
<br /><%
  
		
  end if
  %>
<table width="480" height="25" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr align="center">
    <td height="25" align="center" class="A3"><span class="front8 style3">第<%= page %>页&nbsp;
          <% if page<>1 then %>
          <a href="?page=1<%if request("id")<>"" then%>&id=<%=request("id")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>" class="link2" >首页</a>
          <% else %>
      首页
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?page=<%= page-1 %><%if request("id")<>"" then%>&id=<%=request("id")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>" class="link2" >上一页</a>
      <% else %>
      上一页
      <% end if %>
      &nbsp;
      <select name="select2" onchange='javascript:window.open(this.options[this.selectedIndex].value,&quot;_top&quot;)' class="input2">
        <%For m = 1 To rs1.PageCount%>
        <%if m=page then%>
        <option selected="selected" value="?page=<%=m%><%if request("id")<>"" then%>&id=<%=request("id")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>"><%=m%></option>
        <%else%>
        <option value="?page=<%=m%><%if request("id")<>"" then%>&id=<%=request("id")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>"><%=m%></option>
        <%end if%>
        <% Next %>
      </select>
      <% if page<rs1.pagecount then %>
      <a href="?page=<%= page+1 %><%if request("id")<>"" then%>&id=<%=request("id")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>"  class="link2" >下一页</a>
      <% else %>
      下一页
      <% end if %>
      &nbsp;
      <% if page<rs1.pagecount then %>
      <a href="?page=<%=rs1.pagecount%><%if request("id")<>"" then%>&id=<%=request("id")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>" class="link2" >末页</a>
      <% else %>
      末页
      <% end if %>
      &nbsp;总数<%= rs1.recordcount %>条&nbsp;每页12条</span></td>
  </tr>
</table>
<%
if request("action")="add" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from tp "
		rs1.open sql,conn,1,3
		rs1.addnew
		rs1("pid")=request("pid")
		rs1("tp")=request("image")
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
 response.Redirect "tp.asp?id="&request("pid")
end if
if request("action")="save" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from ct where id="&request("id")
		rs1.open sql,conn,1,3
		rs1("pid")=request("pid")
		rs1("th")=request("th")
		rs1("lx")=request("lx")
		rs1("ftrq")=request("ftrq")
		rs1("msj")=request("msj")
		rs1("wsyh")=request("wsyh")
		rs1("cfd")=request("cfd")
		rs1("syme")=request("syme")
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
 response.Redirect "ctb.asp?id="&request("pid")
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from tp where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
 response.Redirect "tp.asp?id="&request("pid")
end if
%>
</body>
</html>
      


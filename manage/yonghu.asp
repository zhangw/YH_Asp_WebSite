<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="include/db_conn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>获取发布人</title>
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
<table width="480" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  <tr>
    <td width="86" height="25" align="center" bgcolor="#e0e9f8" class="front2">用户名</td>
    <td width="315" align="center" bgcolor="#e0e9f8" class="front2">公司名称</td>
    <td width="95" align="center" bgcolor="#e0e9f8" class="front2">联系人</td>
  </tr>
  <%
  		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from [user] where sh=1 order by id desc"
		xx=0
		rs1.open sql,conn,1,3
		if rs1.eof then
		else
			rs1.pagesize=15
	if request("page")<>"" then 	page=cint(Trim(Request.QueryString("page")))
	if page<1 then
		page=1
	elseif page>rs1.pagecount then
		page=rs1.pagecount
	end if
	rs1.AbsolutePage=page
rowCount = rs1.PageSize
i=0
		do while not rs1.eof and rowcount>0
		i=i+1
	%>	
  <tr onclick="OnClick('<%=rs1("admin")%>')" style="cursor:hand;">
    <td height="25" <%if i mod 2=0 then%>bgcolor="#e0e9f8"<%else%>bgcolor="#ffffff"<%end if%> class="front3">&nbsp;<a href="#" class="link1" onClick="OnClick('<%=rs1("admin")%>')"><%=rs1("admin")%></a></td>
    <td <%if i mod 2=0 then%>bgcolor="#e0e9f8"<%else%>bgcolor="#ffffff"<%end if%> class="front3">&nbsp;<%=rs1("company")%></td>
    <td <%if i mod 2=0 then%>bgcolor="#e0e9f8"<%else%>bgcolor="#ffffff"<%end if%> class="front3">&nbsp;<%=rs1("xm")%></td>
  </tr>
  <%
  
		RowCount=RowCount-1
		rs1.movenext
  loop
  end if
  %>
</table>
<table width="480" height="25" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr align="center">
    <td height="25" align="center" class="A3"><span class="front8 style3">第<%= page %>页&nbsp;
          <% if page<>1 then %>
          <a href="?page=1<%if request("admin")<>"" then%>&admin=<%=request("admin")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>" class="link2" >首页</a>
          <% else %>
      首页
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?page=<%= page-1 %><%if request("admin")<>"" then%>&admin=<%=request("admin")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>" class="link2" >上一页</a>
      <% else %>
      上一页
      <% end if %>
      &nbsp;
      <select name="select2" onchange='javascript:window.open(this.options[this.selectedIndex].value,&quot;_top&quot;)' class="input2">
        <%For m = 1 To rs1.PageCount%>
        <%if m=page then%>
        <option selected="selected" value="?page=<%=m%><%if request("admin")<>"" then%>&admin=<%=request("admin")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>"><%=m%></option>
        <%else%>
        <option value="?page=<%=m%><%if request("admin")<>"" then%>&admin=<%=request("admin")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>"><%=m%></option>
        <%end if%>
        <% Next %>
      </select>
      <% if page<rs1.pagecount then %>
      <a href="?page=<%= page+1 %><%if request("admin")<>"" then%>&admin=<%=request("admin")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>"  class="link2" >下一页</a>
      <% else %>
      下一页
      <% end if %>
      &nbsp;
      <% if page<rs1.pagecount then %>
      <a href="?page=<%=rs1.pagecount%><%if request("admin")<>"" then%>&admin=<%=request("admin")%><%end if%><%if request("type")<>"" then%>&type=<%=request("type")%><%end if%>" class="link2" >末页</a>
      <% else %>
      末页
      <% end if %>
      &nbsp;总数<%= rs1.recordcount %>条&nbsp;每页12条</span></td>
  </tr>
</table>
</body>
</html>
      


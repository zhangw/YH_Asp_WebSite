<!--#include file="include/db_conn.asp"-->
<!--#include file="test.asp"-->
<link href="css/main.css" rel="stylesheet" type="text/css">
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
								<td align="right">您的位置:<A href="index.asp" target="_top">后台管理</A> &gt;&gt; 
									友情链接 &gt;&gt; 类别管理&nbsp;&nbsp;&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>

          <%if request("action")="" or request("action")="list" then%><TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle">信息的增加,删除和修改.</TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
                <TR>
                  <TD height="5"></TD>
                  <TD align="right"></TD>
                </TR>
              </TBODY>
            </TABLE>
		<%
		sql="select * from linkclass"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if  rs.eof then
 response.Write "没有记录!"
else

		%>	
            <TABLE width="90%" border="0" align="center" cellPadding="1" cellSpacing="1" bgColor="#c0c0c0">
              <TBODY>
                <TR bgColor="#e0e9f8">
                  <TD align="middle" width="24" height="24">&nbsp;</TD>
                  <TD align="middle" width="432"><STRONG>标题</STRONG></TD>
                  <TD align="middle" width="50"><STRONG>操 作</STRONG></TD>
                </TR>
<%

		rs.pagesize=12
	if Request.QueryString("page") then page=cint(Trim(Request.QueryString("page")))
	if page<1 then
		page=1
	elseif page>rs.pagecount then
		page=rs.pagecount
	end if
	rs.AbsolutePage=page
rowCount = rs.PageSize
i=0
	do while not rs.eof and rowcount>0
i=i+1
%>
                <TR <%if i mod 2=0 then%>bgColor="#fffff0"<%else%>bgcolor="#F6F6F6"<%end if%>>
                  <TD align="middle" ><%=i%></TD>
                  <TD align="left" ><A href="linkclass.asp?action=edit&id=<%=rs("id")%>"><%=rs("title")%></A></TD>
                  <TD align="middle" ><A title="修改" href="?action=edit&id=<%=rs("id")%>&linkclass=<%=rs("newsclass")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>&linkclass=<%=rs("newsclass")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></TD>
                </TR>
<%
RowCount=RowCount-1
rs.movenext
loop
%>			
              </TBODY>
            </TABLE>
			<table width="90%" border=0 align=center cellSpacing=1 class="navi">  
  <tr>
    <td height="20" ><div align="center" class="unnamed1">第<%= page %>页&nbsp; <a href="?page=1<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh">首页</a> &nbsp;共<%=rs.PageCount%>页&nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >上一页</a>
            <% else %>
        上一页
        <% end if %>
&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=page+1%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >下一页</a>
        <% else %>
        下一页
        <% end if %>
&nbsp;<select name="select" onChange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%For m = 1 To rs.PageCount%>
        <option value="?page=<%=m%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" <%if page=m then%>selected<%end if%>><%=m%></option>
        <% Next %>
      </select>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >末页</a>
        <% else %>
        末页
        <% end if %>
&nbsp;总数<%= rs.recordcount %>条</div></td>
  </tr>
</table>
<%end if%>
            <TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD></TD>
                </TR>
                <TR>
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="?action=add">新增类别</A></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="add" then
%><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 新增类别</td>
            </tr>
            <tr>
              <td></td>
            </tr>
          </tbody>
</table>
<form name="formc" method="post" action="?action=aa"><table width="90%" border=0 align=center cellSpacing=1 class="navi">
    
    <tr>
      <td width="73" height="20" class="front3"><div align="right">类别名称：</div></td>
      <td width="277" class="STYLE5"><input name="title" type="text" id="title" ></td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td height="20" bgcolor="#FFFFFF" class="STYLE5">&nbsp;</td>
      <td bgcolor="#FFFFFF" class="STYLE5"><input type="submit" name="Submit" value="确定"></td>
    </tr>
  </table>
</form>
<%end if
if request("action")="aa" then

	set rs_sel=server.createobject("adodb.recordset")
	sql="select * from linkclass order by id desc" 
	rs_sel.open sql,conn,1,3
	
    set rs=server.createobject("adodb.recordset")
	sql="select * from linkclass " 
	rs.open sql,conn,1,3
	rs.addnew
	rs("title")=request("title")
	rs("newsclass")=rs_sel("newsclass")+1
	rs.update
	rs.close
	

	
	Response.Write "<script>alert('您已经成功添加');location='linkclass.asp'</script>"
end if	
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from linkclass where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 修改类别</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formc" method="post" action="?action=editsave&id=<%=request("id")%>">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    
    <tr>
      <td width="73" height="20" class="front3"><div align="right">类别名称：</div></td>
      <td width="277" class="STYLE5"><input name="title" type="text" id="title" value="<%=rs("title")%>"></td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td height="20" bgcolor="#FFFFFF" class="STYLE5">&nbsp;</td>
      <td bgcolor="#FFFFFF" class="STYLE5"><input type="submit" name="Submit" value="确定"></td>
    </tr>
  </table>
  </form>
<%end if
if request("action")="editsave" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from linkclass where id="&request("id") 
	rs.open sql,conn,1,3
	rs("title")=request("title")
	rs.update
	rs.close
	
	set rs1=server.createobject("adodb.recordset")
	sql="select * from news where linkclass="&request("id") 
	rs1.open sql,conn,1,3
	rs1("title")=request("title")
	rs1.update
	rs1.close
	set rs=nothing
	set rs1=nothing
	Response.Write "<script>alert('您已经成功修改');location='linkclass.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from linkclass where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		
		set rs1=server.createObject("ADODB.Recordset")
		sql="select * from news where linkclass="& request("linkclass") 
		rs1.open sql,conn,2,3
		if not rs1.eof then
		rs1.delete
		rs1.update
		rs1.requery
		end if
		rs1.close
		set rs1=nothing
		
		Response.Write "<script>alert('删除成功');location='linkclass.asp'</script>"
end if
%>		
</body>

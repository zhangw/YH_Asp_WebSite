﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
								<td align="right">您的位置：<A href="index.asp" target="_top">后台管理</A> &gt;&gt; 
									招商加盟</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>

          <%if request("action")="" then%><TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle"> 关于我们的增加、修改和删除。</TD>
                  <TD align="right"></TD>
                </TR>
                <TR>
                  <TD></TD>
                  <TD align="right"></TD>
                </TR>
              </TBODY>
            </TABLE>
			
            <TABLE width="90%" border="0" align="center" cellPadding="1" cellSpacing="1" bgColor="#c0c0c0">
              <TBODY>
                <TR bgColor="#e0e9f8">
                  <TD align="middle" width="24" height="24"></TD>
                  <TD align="middle" bgcolor="#e0e9f8"><strong>标题</strong></TD>
                  <TD align="middle" width="50"><STRONG>操 作</STRONG></TD>
                </TR>
				<%
if request("pid")="" then
sql="select * from zhaos where pid=0 order by id asc"
else
sql="select * from zhaos where pid="&int(request("pid"))&" order by id asc"
end if
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
i=0 
do while not rs.eof
sql="select * from jiameng_class where id="&rs("class")
Set rs1= Server.CreateObject("ADODB.Recordset")
rs1.open sql,conn,1,3
i=i+1
%>
                <TR bgColor="#fffff0">
                  <TD align="middle" bgcolor="#fffff0"><div align="center"><%=i%></div></TD>
                  <TD align="left" bgColor="#fffff0"><%=rs("title")%></TD>
                  <TD align="middle"><A title="修改" href="?action=edit&id=<%=rs("id")%>&pid=<%=request("pid")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>&pid=<%=request("pid")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></TD>
                </TR>
<%
rs.movenext
loop
%>				
              </TBODY>
            </TABLE>
<TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD></TD>
                </TR>
                <TR>

                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="?action=add&pid=<%=int(request("pid"))%>">新增信息</A>&nbsp;<%
sql="select * from zhaos where id="&int(request("pid"))
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if not rs.eof then
%><a href="jiameng.asp?pid=<%=rs("pid")%>">返回上一级</a><%end if%></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="add" then
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 新增信息</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formc" method="post" action="?action=aa&pid=<%=request("pid")%>">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="73" height="20" bgcolor="#FFFFFF"><div align="right"><span class="front3">标题</span>：</div></td>
      <td align="left" bgcolor="#FFFFFF"><input name="title" type="text" id="title">
      <span class="STYLE5"> (标题请不要大于20字)</span></td>
    </tr>
	    <tr bgcolor="#A4B6D7">
      <td height="20" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right"><br />
      内容：</div></td>
      <td height="250" align="left" bgcolor="#FFFFFF" class="STYLE5"><textarea name="theme" style="display:none"></textarea><iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe>	</td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
      <td align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定"></td>
    </tr>
  </table>
</form>
<%end if

if request("action")="aa" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from zhaos "
	rs.open sql,conn,1,3
	rs.addnew()
	rs("title")=request("title")
	rs("class")=request("class")
	rs("theme")=request("theme")
	rs("pid")=int(request("pid"))
	rs.update
	rs.close
	set rs=nothing		
	Response.Write "<script>alert('您已经成功添加');location='zs.asp?pid="&request("pid")&"'</script>"
end if	
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from zhaos where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 修改信息</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formc" method="post" action="?action=editsave&pid=<%=request("pid")%>&id=<%=request("id")%>">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="73" height="20" bgcolor="#FFFFFF"><div align="right"><span class="front3">标题</span>：</div></td>
      <td align="left" bgcolor="#FFFFFF"><input name="title" type="text" id="title" value="<%=rs("title")%>">
      <span class="STYLE5"> (标题请不要大于20字)</span></td>
    </tr>
	    <tr bgcolor="#A4B6D7">
      <td height="20" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right"><br />
      内容：</div></td>
      <td height="250" align="left" bgcolor="#FFFFFF" class="STYLE5"><textarea name="theme" style="display:none"><%=rs("theme")%></textarea><iframe id="editor2" src="../pic/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe>	</td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
      <td align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定"></td>
    </tr>
  </table>
</form>
<%end if

if request("action")="editsave" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from zhaos where id=" & request("id") 
	rs.open sql,conn,1,3
	rs("title")=request("title")
	rs("class")=request("class")
	rs("theme")=request("theme")
	rs("pid")=int(request("pid"))
	rs.update
	rs.close
	set rs=nothing	
	Response.Write "<script>alert('您已经成功修改');location='zs.asp?pid="&request("pid")&"'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from zhaos where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('删除成功');location='zs.asp?pid="&request("pid")&"'</script>"
end if
%>		
</body>
</html>

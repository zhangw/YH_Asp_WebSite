<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title></title>
<!--#include file="include/db_conn.asp"-->
<!--#include file="test.asp"-->
<%
Response.Expires = 0
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<link href="css/main.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
       Sub MainFl(celclass) 
           Dim Rs 
           Set Rs = Conn.Execute("SELECT id,class FROM productclass WHERE father=0") 
           If Not Rs.Eof Then 
               Do While Not Rs.Eof 
			       if celclass=Rs("id") then
                   response.Write "<option value=""" & Trim(Rs("id")) & """ selected>" & Trim(Rs("class")) & "</option>" 
				   else
				   response.Write "<option value=""" & Trim(Rs("id")) & """>" & Trim(Rs("class")) & "</option>" 
				   end if
                   Call Subfl(Rs("id"),Rs("class")&"→",celclass) '循环子级分类 
                   response.Write "</div>" 
               Rs.MoveNext 
               If Rs.Eof Then Exit Do '防上造成死循环 
               Loop 
           End If 
           Set Rs = Nothing 
       End Sub 
%>
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
									设备展示&gt;&gt;设备管理</td>
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
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle">产品 信息的增加、修改和删除。</TD>
                  <TD align="right"></TD>
                </TR>
                <TR>
                  <TD></TD>
                  <TD align="right"></TD>
                </TR>
              </TBODY>
            </TABLE>
			
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="164" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">信息标题</div></th>
    <th width="146" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">类别</div></th>
    <th width="94" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">操作</div></th>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from wzjs " 
	sql=sql & " order by id desc"
	rs.open sql,conn,1,3
if  rs.eof then
 response.Write "没有记录!"
 response.End()
end if
		rs.pagesize=13
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
	set rs1=server.createobject("adodb.recordset")
	sql="select * from server_class where id="&rs("newsclass")
	rs1.open sql,conn,1,3 
%>
  <tr <%if i mod 2=0 then%>bgcolor="#fffff0"<%else%>bgcolor="#ffffff"<%end if%> class="unnamed1">
    <td width="164" height="20"  class="STYLE5" style="padding-left:5px;"><%if len(rs("title"))>13 then%><%=mid(rs("title"),1,13)%>...<%else%><%=rs("title")%><%end if%></td>
	<td width="146"  class="STYLE5"><div align="center"><%=rs1("title")%></div></td>
    <td width="94" ><div align="center"><A title="修改" href="?action=edit&id=<%=rs("id")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></div></td>
  </tr>
  <%RowCount=RowCount-1
rs.movenext
loop
%>
</table>
<table width="90%" border=0 align=center cellSpacing=1 class="navi">  
  <tr>
    <td height="20" ><div align="center" class="unnamed1">第<%= page %>页&nbsp; <a href="?page=1<%if request("title")<>"" then%>&title=<%=request("title")%><%end if%><%if request("theme")<>"" then%>&theme=<%=request("theme")%><%end if%><%if request("lb")<>"" then%>&lb=<%=request("lb")%><%end if%>" class="hh">首页</a> &nbsp;共<%=rs.PageCount%>页&nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %><%if request("title")<>"" then%>&title=<%=request("title")%><%end if%><%if request("theme")<>"" then%>&theme=<%=request("theme")%><%end if%><%if request("lb")<>"" then%>&lb=<%=request("lb")%><%end if%>" class="hh" >上一页</a>
            <% else %>
        上一页
        <% end if %>
&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%><%if request("title")<>"" then%>&title=<%=request("title")%><%end if%><%if request("theme")<>"" then%>&theme=<%=request("theme")%><%end if%><%if request("lb")<>"" then%>&lb=<%=request("lb")%><%end if%>" class="hh" >下一页</a>
        <% else %>
        下一页
        <% end if %>
&nbsp;<select name="select" onChange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%For m = 1 To rs.PageCount%>
        <option value="?page=<%=m%><%if request("title")<>"" then%>&title=<%=request("title")%><%end if%><%if request("theme")<>"" then%>&theme=<%=request("theme")%><%end if%><%if request("lb")<>"" then%>&lb=<%=request("lb")%><%end if%>" <%if page=m then%>selected<%end if%>><%=m%></option>
        <% Next %>
      </select>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%><%if request("title")<>"" then%>&title=<%=request("title")%><%end if%><%if request("theme")<>"" then%>&theme=<%=request("theme")%><%end if%><%if request("lb")<>"" then%>&lb=<%=request("lb")%><%end if%>" class="hh" >末页</a>
        <% else %>
        末页
        <% end if %>
&nbsp;总数<%= rs.recordcount %>条</div></td>
  </tr>
</table>

            <TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD></TD>
                </TR>
                <TR>
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="sb_add.asp">新增信息</A></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from wzjs where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 修改设备信息</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="100"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">设备名称：</div></td>
      <td width="596" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;" value="<%=rs("title")%>"></td>
    </tr>
     <tr>
      <td width="100" height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">类&nbsp;&nbsp;&nbsp;&nbsp;别：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="class">
<%	      set rs1=server.createobject("adodb.recordset")
	sql="select * from server_class" 
	rs1.open sql,conn,1,3
	do while not rs1.eof
	 if rs1("id")=rs("newsclass") then
%>
<option value="<%=rs1("id")%>" selected><%=rs1("title")%></option>
<%	 
	 else
%>
        <option value="<%=rs1("id")%>"><%=rs1("title")%></option>
<%
end if
rs1.movenext
loop
%> 
        </select></td>
    </tr>
    <tr>
      <td width="100" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">设备说明：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"><%=rs("content")%></textarea>
	                <iframe id="editor2" src="../editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe></td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
      <th height="20" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定" /></th>
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from wzjs where id="&request("id")
		rs1.open sql,conn,1,3
		rs1("title")=request("title")
		rs1("content")=request("theme")
		rs1("newsclass")=request("class")
		'response.Write request("class")
		'response.End()
		rs1("time")=now()
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		Response.Write "<script>alert('您已经成功修改');location='sb_list.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from wzjs where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('删除成功');location='sb_list.asp'</script>"
end if
%>		
</body>
</html>

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
								<td align="right">您的位置：<A href="index.asp" target="_top">后台管理</A> &gt;&gt; 
									留言管理</td>
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
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle">设计咨询信息的回复和删除。</TD>
                  <TD align="right"></TD>
                </TR>
                <TR>
                  <TD></TD>
                  <TD align="right"></TD>
                </TR>
              </TBODY>
            </TABLE>
	<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#A4B6D7" class="unnamed1">
    <td width="84" height="20" bgcolor="#e0e9f8" class="STYLE5"><div align="center"><strong>姓名</strong></div></td>
    <td width="119" bgcolor="#e0e9f8" class="STYLE5"><div align="center"><strong>时间</strong></div></td>
    <td bgcolor="#e0e9f8" class="STYLE5"><div align="center"><strong>内容</strong></div>     </td>
    <td width="92" height="25" bgcolor="#e0e9f8" class="STYLE5"><div align="center"><strong>操作</strong></div></td>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from contact " 
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
%>
  <tr  class="unnamed1" bgcolor="#ffffff"  onMouseOut="this.style.backgroundColor='#ffffff'" onMouseOver="this.style.backgroundColor='#F0E27D'">
    <td width="85" height="25" align="center" class="STYLE5" ><%=rs("name")%></td>
	<td width="120"  class="STYLE5"><div align="center"><%=rs("time")%></div></td>
    <td align="left" class="STYLE5"  ><%if len(rs("content"))>20 then%><%=mid(rs("title"),1,20)%>...<%else%><%=rs("content")%><%end if%></td>
    <td width="93"><div align="center"><a href="?action=edit&id=<%=rs("id")%><%if request("lb")<>"" then%>&lb=<%=request("lb")%><%end if%>" class="hh">查看</a>|<a href="?action=del&id=<%=rs("id")%>" class="hh">删除</a></div></td>
  </tr>
  <%RowCount=RowCount-1
  i=i+1
rs.movenext
loop
%>
</table>
<table width="90%"  border="0" align="center" cellspacing="1">
  <tr>
    <td height="20"><div align="center" class="unnamed1">第<%= page %>页&nbsp; 首页 &nbsp;共<%=rs.PageCount%>页&nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %>&action=list<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >上一页</a>
            <% else %>
        上一页
        <% end if %>
&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%>&action=list<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >下一页</a>
        <% else %>
        下一页
        <% end if %>
&nbsp;
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%>&action=list<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >末页</a>
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
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
sql="select * from contact where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" />查看留言</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<script type="text/javascript">
	function check(){
		if(document.formm.reform.value == ""){
			alert("请输入回复内容?");
			return false;
		}
		return true;
	}
</script>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>" onsubmit="return check()">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
    
    <tr>
      <td width="157" height="25" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">姓名：</div></td>
      <td width="216" bgcolor="#FFFFFF">&nbsp;<%=rs("name")%></td>
      <td width="897" bgcolor="#FFFFFF">发表时间：&nbsp;<%=rs("time")%></td>
    </tr>
    <tr>
      <td width="157" height="25" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">电话：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">&nbsp;<%=rs("phone")%></td>
    </tr>
    <tr>
      <td width="157" height="25" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">电子邮件：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">&nbsp;<%=rs("e_mail")%></td>
    </tr>
    <tr>
      <td width="157" height="25" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">内容：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">&nbsp;<%=rs("content")%></td>
    </tr>
    <tr>
      <td width="157" height="25" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">回复：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><textarea name="reform" rows="4" id="reform" style="width:300px;"><%=rs("reform")%></textarea></td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td width="157" height="25" align="right" bgcolor="#FFFFFF"><div align="right">是否显示：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="xs" type="checkbox" id="xs" value="1" <%if rs("show") then%>checked<%end if%>>是&nbsp;</td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td width="157" height="25" align="right" bgcolor="#FFFFFF">&nbsp;</td>
      <td colspan="2" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定"></td>
    </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
	if request("xs")="1" then
	i=1
	else
	i=0
	end if
	conn.execute "update contact set reform='"&request("reform")&"',show="&i&" where id="&request("id")
	Response.Write "<script>alert('您已经成功修改');location='book.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from contact where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('删除成功');location='book.asp'</script>"
end if
%>		
</body>
</html>

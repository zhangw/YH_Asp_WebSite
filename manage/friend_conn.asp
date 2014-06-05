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
									信息发布&gt;&gt;信息管理</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>

          <%if request("action") = "list" then%><TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle"> 信息的增加、修改和删除。</TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
                <TR>
                  <TD>&nbsp;</TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="244" height="20" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">序号</div></th>
    <th width="211" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">标题</div></th>
    <th bgcolor="#e0e9f8"  class="STYLE5"><div align="center">发布时间</div></th>
    <th width="127" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">操作</div></th>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
			sql="select * from firend_conn " 
			sql=sql & " order by id desc"
			rs.open sql,conn,1,3
		if  rs.eof then
 			response.Write "没有记录!"
		else
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
		do while not rs.eof
			i=i+1
%>
  <tr <%if RowCount mod 2=0 then%>bgColor="#fffff0"<%else%>bgcolor="#F6F6F6"<%end if%> class="unnamed1">
    <td width="244" height="20"  class="STYLE5" style="padding-left:5px;"><%=i%></td>
	<td width="211"  class="STYLE5"><div align="center"><span class="STYLE5" style="padding-left:5px;">
	  <%if len(rs("title"))>13 then%>
	  <%=mid(rs("title"),1,13)%>...
	  <%else%>
	  <%=rs("title")%>
	  <%end if%>
	</span></div></td>
    <td  class="STYLE5"><div align="center"><%=formatdatetime(rs("time"),2)%></div></td>
    <td width="127" ><div align="center"><A title="修改" href="?action=edit&id=<%=rs("id")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></div></td>
  </tr>
	<%
		rs.movenext
		loop
		end if
	%>
</table>
<table width="90%" border=0 align=center cellSpacing=1 class="navi">  
  <tr>
    <td height="20" ><div align="center" class="unnamed1">第<%= page %>页&nbsp; 
	<% if page > request.QueryString("page") then%>
	<a href="?page=1" class="hh">首页</a> 
	<%else%>
		首页
	<%end if%>
	&nbsp;共<%=rs.PageCount%>页&nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %>" class="hh" >上一页</a>
            <% else %>
        		上一页
       		<% end if %>
	&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=page+1%>" class="hh" >下一页</a>
        <% else %>
        	下一页
        <% end if %>
&nbsp;<select name="select" onChange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%For m = 1 To rs.PageCount%>
        <option value="?page=<%=m%>" <%if page=m then%>selected<%end if%>><%=m%></option>
        <% Next %>
      </select>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%>" class="hh" >末页</a>
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
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> 
				  <A href="friend_conn.asp?action=add">新增信息</A></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from friend_conn where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 编辑信息</td>
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
      <td width="100"  height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">标&nbsp;&nbsp;&nbsp;&nbsp;题：</div></td>
      <td width="596" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;" value="<%=rs("title")%>"></td>
    </tr>
    <tr>
      <td width="100" height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">地&nbsp;&nbsp;&nbsp;&nbsp;址：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="address" type="text" id="address" value="<%=rs("link")%>" size="50" /></td>
    </tr>
    <input name="image" type="hidden" id="image"/>
    <tr>
      <td width="100" height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">上传图片：</div></td>
      <td height="20" bgcolor="#FFFFFF" class="STYLE5"><iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25" style="height:20px;"></iframe></td>
    </tr>	
    <tr bgcolor="#A4B6D7">
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
      <th height="20" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定" /></th>
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%end if%>
<% if request("action") = "add" then%>
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 添加信息</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
	<form name="formm" method="post" action="?action=aa" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="100"  height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">标&nbsp;&nbsp;&nbsp;&nbsp;题：</div></td>
      <td width="596" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;" ></td>
    </tr>
    <tr>
      <td width="100" height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">地&nbsp;&nbsp;&nbsp;&nbsp;址：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="address" type="text" id="address" size="50" /></td>
    </tr>
    <input name="image" type="hidden" id="image"/>
    <tr>
      <td width="100" height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">上传图片：</div></td>
      <td height="20" bgcolor="#FFFFFF" class="STYLE5"><iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
    </tr>			
    <tr bgcolor="#A4B6D7">
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
      <th height="20" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定" /></th>
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%end if%>
<script>
	function check(){
		if(document.formm.title.value == ""){
			alert("标题不能为空?");
			return false;
		}
		if(document.formm.address.value == ""){
			alert("地址不能为空?");
			return false;
		}
		if(document.formm.image.value == ""){
			alert("请上传图片?");
			return false;
		}
		return true;
	}
</script>
<%
if request("action")="aa" then
	title = trim(request("title"))
	address = trim(request("address"))
	pic = trim(request("image"))
	set rs = server.CreateObject("adodb.recordset")
		sql = "select * from friend_conn"
		rs.open sql,conn,1,3
		rs.addnew
		rs(1) = title
		rs(2) = address
		rs(3) = pic
		rs.update
		rs.close
		set rs = nothing
		response.Write "<script>alert('您已经成功添加');location.reload('friend_conn.asp?action=list')</script>"
end if
if request("action")="editsave" then
    title = trim(request("title"))
	address = trim(request("address"))
	pic = trim(request("image"))
	set rs = server.CreateObject("adodb.recordset")
		sql = "select * from friend_conn where id="&request.QueryString("id")
		rs.open sql,conn,1,3
		rs(1) = title
		rs(2) = address
		rs(3) = pic
		rs.update
		rs.close
		set rs = nothing
	Response.Write "<script>alert('您已经成功修改');location.reload('friend_conn.asp?action=list')</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from friend where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('删除成功');location.reload('friend_conn.asp?action=list')</script>"
end if
%>		
</body>
</html>

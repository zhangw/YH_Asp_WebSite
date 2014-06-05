<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title></title>
<!--#include file="include/db_conn.asp"-->
<!--#include file="test.asp"-->
<!--#include file="../inc/str.asp"-->
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

          <%if request("action")="" then%><TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
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
			
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><table width="186" border="0" cellspacing="0" cellpadding="0">
               <form id="form1" name="form1" method="post" action=""> <tr>
                  <td width="146"><input name="titles" type="text" id="titles" style="height:15px;"/></td>
                 <td width="40"><input type="submit" name="Submit2" value="查找" />                 </td>
                </tr> </form>
              </table></td>
            </tr>
          </table>
          <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
 <form id="form2" name="form2" method="post" action="?action=del&newsid=<%=rs("newsid")%>&page=<%=int(request("page"))%>"> <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="28" bgcolor="#e0e9f8"  class="STYLE5">&nbsp;</th>
    <th width="457" height="20" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">标题</div></th>
    <th width="226" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">发布人</div></th>
    <th width="196" bgcolor="#e0e9f8"  class="STYLE5">分类</th>
    <th width="255" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">发布时间</div></th>
    <th width="102" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">操作</div></th>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from news where title<>'' " 
	    if request("titles")<>"" then
		sql=sql&" and title like '%" & request("titles") & "%'"
		end if
		if request("newsid")<>"" then
				sql=sql&" and newsid="&request("newsid")
		end if

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
	do while not rs.eof and rowcount>0

%>

  <tr bgcolor="#ffffff" class="unnamed1" onMouseOut="this.style.backgroundColor='#ffffff'" onMouseOver="this.style.backgroundColor='#F0E27D'">
    <td width="28"   class="STYLE5" style="padding-left:5px;">
      <input name="id" type="checkbox" id="id" value="<%=rs("id")%>" />    </td>
    <td width="457" height="20"  class="STYLE5" style="padding-left:5px;"><%if len(rs("title"))>32 then%><%=mid(rs("title"),1,32)%>...<%else%><%=rs("title")%><%end if%></td>
	<%
	set rs1=server.createobject("adodb.recordset")
	sql="select * from newsid where id="&rs("newsid")
	rs1.open sql,conn,1,3 
    %>
	<td width="226"   class="STYLE5"><div align="center"><%=rs("fabu")%></div></td>
    <td align="center"   class="STYLE5"><%=rs1("title")%></td>
    <td   class="STYLE5"><div align="center"><%=formatdatetime(rs("time"),2)%></div></td>
    <td width="102"  ><div align="center"><A title="修改" href="?action=edit&id=<%=rs("id")%>&title=<%=rs("title")%>&page=<%=int(request("page"))%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>&page=<%=int(request("page"))%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></div></td>
  </tr>
  <%RowCount=RowCount-1
rs.movenext
loop
%>
  <tr bgcolor="#ffffff" class="unnamed1" onMouseOut="this.style.backgroundColor='#ffffff'" onMouseOver="this.style.backgroundColor='#F0E27D'">
    <td height="20" colspan="6" align="right"  class="STYLE5" style="padding-left:5px;"><a href="#" onclick="cc('form2',1)">全选</a> <a href="#" onclick="cc('form2',0)">全不选</a> <a href="#" onclick="cc('form2',2)">反选</a>&nbsp;
      <input type="submit" name="Submit3" value="删除选中" /></td>
    </tr>
</form>
</table>
<script language="Javascript">  
function cc(N,bool){  
  var aa = document.getElementById(N).getElementsByTagName("input");
  for (var i=0; i<aa.length; i++){
	  if (aa[i].type=="checkbox")
		aa[i].checked = bool==1 ? true : (bool==0 ? false : !aa[i].checked);
  }  
}
</script>
<table width="90%" border=0 align=center cellSpacing=1 class="navi">  
  <tr>
    <td height="20" ><div align="center" class="unnamed1">第<%= page %>页&nbsp; <a href="?page=1<%if request("title")<>"" then%>&title=<%=request("title")%><%end if%><%if request("newsid")<>"" then%>&newsid=<%=request("newsid")%><%end if%>" class="hh">首页</a> &nbsp;共<%=rs.PageCount%>页&nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %><%if request("titles")<>"" then%>&titles=<%=request("titles")%><%end if%><%if request("newsid")<>"" then%>&newsid=<%=request("newsid")%><%end if%>" class="hh" >上一页</a>
            <% else %>
        上一页
        <% end if %>
&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=page+1%><%if request("titles")<>"" then%>&titles=<%=request("titles")%><%end if%><%if request("newsid")<>"" then%>&newsid=<%=request("newsid")%><%end if%>" class="hh" >下一页</a>
        <% else %>
        下一页
        <% end if %>
&nbsp;<select name="select" onChange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%For m = 1 To rs.PageCount%>
        <option value="?page=<%=m%><%if request("titles")<>"" then%>&titles=<%=request("titles")%><%end if%><%if request("newsid")<>"" then%>&newsid=<%=request("newsid")%><%end if%>" <%if page=m then%>selected<%end if%>><%=m%></option>
        <% Next %>
      </select>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%><%if request("titles")<>"" then%>&titles=<%=request("titles")%><%end if%><%if request("newsid")<>"" then%>&newsid=<%=request("newsid")%><%end if%>" class="hh" >末页</a>
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
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="news_add.asp">新增信息</A></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from news where id="& request("id")
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
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>&page=<%=int(request("page"))%><%if request("titles")<>"" then%>&titles=<%=request("titles")%><%end if%><%if request("newsid")<>"" then%>&newsid=<%=request("newsid")%><%end if%>">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="150"  height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">标&nbsp;&nbsp;&nbsp;&nbsp;题：</div></td>
      <td width="596" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:550px;" value="<%=rs("title")%>"></td>
    </tr>
	    <tr>
      <td height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">类&nbsp;&nbsp;&nbsp;&nbsp;别：</div></td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><select name="newsids" id="newsids">
	  <%
	set rs1=server.createobject("adodb.recordset")
	sql="select * from newsid  order by id asc" 

	rs1.open sql,conn,1,1
	do while not rs1.eof
	%>
        <option value="<%=rs1("id")%>" <%if rs("newsid")=rs1("id") then%>selected="selected"<%end if%>><%=rs1("title")%></option>
	<%
	
	rs1.movenext
	loop
	%>
      </select></td>
    </tr>
	    <tr>
      <td width="150" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">关 键 字：</td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><input name="gjz" type="text" id="gjz" style="width:550px;" value="<%=rs("gjz")%>"/></td>
    </tr>
	    <tr>
      <td width="150" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">描&nbsp;&nbsp;&nbsp;&nbsp;述：</td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><textarea name="ms" id="ms" style="width:550px;"><%=rs("ms")%></textarea></td>
    </tr>
	    <tr>
      <td width="150" height="20" align="center" valign="middle" bgcolor="#FFFFFF" class="STYLE5"><div align="right">发 布 人：</div></td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><input name="fabu" type="text" id="fabu" style="width:550px;" value="<%=rs("fabu")%>" /></td>
    </tr>
    <tr>
      <td width="150" height="20" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">内&nbsp;&nbsp;&nbsp;&nbsp;容：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"><%=rs("content")%></textarea>
	                <iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="560" height="405"></iframe></td>
    </tr>
		    <input name="image" type="hidden" id="image" value="<%=rs("pic")%>"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">上传文件图片<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe> </td>
    </tr>	
    <tr bgcolor="#A4B6D7">
      <th width="150" height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <div align="right"></div></th>
      <th height="20" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定修改" /></th>
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
    theme=trim(request("theme"))
    set rs=server.createobject("adodb.recordset")
	sql="select * from news where id="&request("id") 
	rs.open sql,conn,1,3
	rs("title")=trim(request("title"))
	rs("gjz")=request("gjz")
	rs("ms")=request("ms")
	rs("city")=request("city")
	rs("content")=theme
	rs("newsid")=request("newsids")
	rs("pic")=request("image")
    rs("zxdt")=int(request("zxdt"))
	rs("tthg")=int(request("tthg"))
	rs("tpxw")=int(request("tpxw"))
	rs("fabu")=request("fabu")
	rs.update
	rs.close
	set rs=nothing	
	

	Response.Write "<script>alert('您已经成功修改');location='news_list.asp?page="&int(request("page"))&"&titles="&request("titles")&"&newsid="&request("newsid")&"'</script>"
end if
if request("action")="del" then
		sql="DELETE  from news where id in("& request("id")&")" 
		conn.execute(sql)
	response.Redirect "news_list.asp?page="&request("page")&"&titles="&request("titles")&"&newsid="&request("newsid")
end if
function delfile(filepath)
imangepath=trim(filepath)
path=server.MapPath(imangepath)
SET fs=server.CreateObject("Scripting.FileSystemObject")
if FS.FileExists(path) then
FS.DeleteFile(path) 
end if 
set fs=nothing
end function 
%>		
</body>
</html>

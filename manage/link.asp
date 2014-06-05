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
       '定义子级分类 
       Sub SubFl(FID,StrDis,celclass) 
           Dim Rs1 
		   aa=StrDis
           Set Rs1 = Conn.Execute("SELECT id,class FROM productclass WHERE father = " & FID & "")
           If Not Rs1.Eof Then 
               Do While Not Rs1.Eof 
			       if celclass=Rs1("id") then
				   response.Write "<option value="""& Trim(Rs1("id")) & """ selected>" & StrDis & Trim(Rs1("class")) & "</option>"
				   else
                   response.Write "<option value="""& Trim(Rs1("id")) & """>" & StrDis & Trim(Rs1("class")) & "</option>" 
                   end if
				   Call SubFl(Trim(Rs1("id")),StrDis&Rs1("class")&"→",celclass) '递归子级分类 
               Rs1.Movenext:Loop 
               If Rs1.Eof Then 
                   Rs1.Close 
                   Exit Sub 
               End If 
           End If 
           Set Rs1 = Nothing 
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
								<td align="right">您的位置：<A href="default.htm" target="_top">后台管理</A> &gt;&gt; 
									订单管理&gt;&gt;订单管理</td>
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
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle">友情链接的增加、修改和删除。</TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
                <TR>
                  <TD>&nbsp;</TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
<%
     set rs=server.createobject("adodb.recordset")
	sql="select * from link order by id desc" 
	rs.open sql,conn,1,3
if  rs.eof then
 response.Write "没有记录!"
else
%>			
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="321" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">标题</div></th>
    <th width="104" bgcolor="#e0e9f8"  class="STYLE5">类别</th>
    <th width="733" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">链接地址</div></th>
    <th width="84" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">操作</div></th>
  </tr>
  <%
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
sql="select * from linkclass where id="&rs("class")
Set rs1= Server.CreateObject("ADODB.Recordset")
rs1.open sql,conn,1,3
%>
  <tr <%if i mod 2=0 then%>bgcolor="#fffff0"<%else%>bgcolor="#ffffff"<%end if%> class="unnamed1">
    <td width="321" height="25"  class="STYLE5" style="padding-left:5px;"><a href="../prolist.asp?id=<%=rs("id")%>" target="_blank"><%if len(rs("title"))>13 then%><%=mid(rs("title"),1,13)%>...<%else%><%=rs("title")%><%end if%></a></td>
	<td align="center"  class="STYLE5"><%=rs1("title")%></td>
	<td  class="STYLE5"><div align="center"><%=rs("link")%></div></td>
    <td width="84" ><div align="center"><A title="修改" href="?action=edit&id=<%=rs("id")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></div></td>
  </tr>
  <%
i=i+1  
RowCount=RowCount-1
rs.movenext
loop
%>
</table>
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
        <a href="?page=<%=rs.pagecount%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >下一页</a>
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
                  <TD>&nbsp;</TD>
                </TR>
                <TR>
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="?action=add">新增信息</A></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from link where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><p><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 修改信息</p>
              </td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>">
  <table width="90%" border=0 align=center cellspacing=1 class="navi">
    <tr>
      <td width="153"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">类别：</td>
      <td height="20" colspan="2" bgcolor="#FFFFFF"><select name="class" id="class">
<%
sql="select * from linkclass "
Set rs1= Server.CreateObject("ADODB.Recordset")
rs1.open sql,conn,1,3
do while not rs1.eof
%>      
	    <option value="<%=rs1("id")%>" <%if rs("class")=rs1("id") then%> selected="selected"<%end if%>><%=rs1("title")%></option>
<%
rs1.movenext
loop
%>
      </select>
      </td>
    </tr>
    <tr>
      <td  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">标题：</div></td>
      <td height="20" colspan="2" bgcolor="#FFFFFF"><span class="front3">
        <input name="title" type="text" id="title"  value="<%=rs("title")%>"/>
      </span></td>
    </tr>
    <tr>
      <td height="20" align="right" bgcolor="#FFFFFF" class="front3">链接<span class="STYLE5">：</span></td>
      <td width="921" height="20" bgcolor="#FFFFFF" class="front3"><input name="link" type="text" id="link"  value="<%=rs("link")%>"/></td>
    </tr>
		  <!--  <input name="image" type="hidden" id="image" value="<%=rs("pic")%>"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">上传小图<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe> 
      (145*67)</td>
    </tr>	 -->
	
	  <tr> 
                  <td align="right">缩略图： 
                    <input name="image" type="hidden" id="IncludePic" value="<%=rs("pic")%>"></td>
                  <td height="27"> <input name="defaultpicurl" type="text" id="defaultpicurl" value="img/nopic.gif" size="40" maxlength="120"><input name="UploadFiles" type="hidden" id="UploadFiles2">(请复制上传后的地址填入)</td>
                </tr>
                
                  <tr> 
                    <td height="42"> <div align="right"></div></td>
                    <td><iframe src="upfileother1.asp?atype=defaultpicurl" frameborder=0 width="63%" height="21" scrolling=no name="1"></iframe></td>
                  </tr> 
    <tr bgcolor="#A4B6D7">
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
      <th height="20" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit2" value="确定" /></th>
      <th width="171" height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%end if
if request("action")="add" then%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><p><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 添加信息</p>
              </td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formm" method="post" action="?action=addsave&id=<%=request("id")%>">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
        <tr>
      <td width="152"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">类别：</td>
      <td height="20" colspan="2" bgcolor="#FFFFFF"><select name="class" id="class">
<%
sql="select * from linkclass "
Set rs1= Server.CreateObject("ADODB.Recordset")
rs1.open sql,conn,1,3
do while not rs1.eof
%>      
	    <option value="<%=rs1("id")%>" ><%=rs1("title")%></option>
<%
rs1.movenext
loop
%>
      </select>
      </td>
    </tr>
	<tr>
      <td  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">标题：</div></td>
      <td height="20" colspan="2" bgcolor="#FFFFFF"><span class="front3">
        <input name="title" type="text" id="title"/>
      </span></td>
    </tr>
    
	    <tr>
      <td height="20" align="right" bgcolor="#FFFFFF" class="front3">链接<span class="STYLE5">：</span></td>
      <td width="922" height="20" bgcolor="#FFFFFF" class="front3"><input name="link" type="text" id="link" value="http://"/></td>
    </tr>
		    <input name="image" type="hidden" id="image"/>
 <!--   <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">上传小图<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe> 
      (145*67)</td>
    </tr>	-->
	  <tr> 
                  <td align="right">缩略图： 
                    <input name="IncludePic" type="hidden" id="IncludePic" value="yes"></td>
                  <td height="27"> <input name="defaultpicurl" type="text" id="defaultpicurl" value="img/nopic.gif" size="40" maxlength="120"><input name="UploadFiles" type="hidden" id="UploadFiles2">(请复制上传后的地址填入)</td>
                </tr>
                
                  <tr> 
                    <td height="42"> <div align="right"></div></td>
                    <td><iframe src="upfileother1.asp?atype=defaultpicurl" frameborder=0 width="63%" height="21" scrolling=no name="1"></iframe></td>
                  </tr> 
    <tr bgcolor="#A4B6D7">
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
      <th height="20" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定" /></th>
      <th width="171" height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from link where id="&request("id")
		rs1.open sql,conn,1,3
		rs1("title")=request("title")
		rs1("link")=request("link")
		rs1("pic")=request("image")
		rs1("class")=request("class")
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		Response.Write "<script>alert('您已经成功修改');location='link.asp'</script>"
end if
if request("action")="addsave" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from link"
		rs1.open sql,conn,1,3
		rs1.addnew
		rs1("title")=request("title")
		rs1("link")=request("link")
		rs1("pic")=request("image")
		rs1("class")=request("class")
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		Response.Write "<script>alert('您已经成功添加');location='link.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from link where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('删除成功');location='link.asp'</script>"
end if
%>		
</body>
</html>

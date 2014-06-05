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
								<td align="right">您的位置：<A href="index.asp" target="_top">后台管理</A> &nbsp;&nbsp;</td>
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
                  <TD height="21"><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle"> 信息的增加、修改和删除。</TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
                <TR>
                  <TD><table width="300" border="0" cellspacing="0" cellpadding="0">
               <form id="form1" name="form1" method="post" action=""> <tr>
                  <td width="146"><input name="titles" type="text" id="titles" style="height:15px;"/></td>
                 <td width="114">
                   <select name="newsid" id="newsid">
<%
sql="select  * from product_class1  order by id asc"
set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql, conn, 1, 1
do while not rs1.eof
%>   
					  <option value="<%=rs1("id")%>"><%=rs1("classname")%></option>
<%
rs1.movenext
loop
%>
                    </select>
                 
                 </td>
                 <td width="40"><input type="submit" name="Submit2" value="查找" />
                 </td>
                </tr> </form>
              </table></TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
	<script language="Javascript">  
function cc(N,bool){  
  var aa = document.getElementById(N).getElementsByTagName("input");
  for (var i=0; i<aa.length; i++){
	  if (aa[i].type=="checkbox")
		aa[i].checked = bool==1 ? true : (bool==0 ? false : !aa[i].checked);
  }  
}
</script>		
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
<form id="form2" name="form2" method="post" action="?action=del&newsid=<%=request("newsid")%>&titles=<%=request("titles")%>&page=<%=int(request("page"))%>">
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="24" bgcolor="#e0e9f8"  class="STYLE5">&nbsp;</th>
	<th width="577" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">标题</div></th>
    <th width="284" bgcolor="#e0e9f8"  class="STYLE5">产品类型</th>
    <th width="284" bgcolor="#e0e9f8"  class="STYLE5">发布时间</th>
	 <th width="108" bgcolor="#e0e9f8"  class="STYLE5">排序</th> 
    <th width="195" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">操作</div></th>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from product1 where title<>'' "
	if request("titles")<>"" then 
	sql=sql&" and title like '%"&request("titles")&"%'"
	end if
	if request("newsid")<>"" then 
	sql=sql&" and productclass1 ="&request("newsid")
	end if
	sql=sql & " order by rootid asc"
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
		set rs1=server.createObject("ADODB.Recordset")
	sql="select * from product_class1 where id="& rs("productclass1")
	rs1.open sql,conn,2,3
%>
  <tr bgcolor="#ffffff"  onMouseOut="this.style.backgroundColor='#ffffff'" onMouseOver="this.style.backgroundColor='#F0E27D'" class="unnamed1">
   <td width="24"   class="STYLE5" style="padding-left:5px;">
      <input name="id" type="checkbox" id="id" value="<%=rs("id")%>" />    </td>
    <td height="20"  class="STYLE5" style="padding-left:5px;"><a href="?action=edit&amp;id=<%=rs("id")%>&amp;BigClassName=<%=request("BigClassName")%>">
      <%if len(rs("title"))>20 then%>
      <%=mid(rs("title"),1,20)%>...
      <%else%>
      <%=rs("title")%>
      <%end if%>
    </a></td>
    <td width="284" align="center"  class="STYLE5"><%=rs1("classname")%></td>
    <td width="284" align="center"  class="STYLE5"><%=rs("time")%></td>
	 <td width="108" align="center" ><a href="?action=sy&rootid=<%=rs("rootid")%>&page=<%=request("page")%>">上移</a>&nbsp;<a href="?action=xy&rootid=<%=rs("rootid")%>&page=<%=request("page")%>">下移</a></td>
    <td width="195" ><div align="center"><A title="修改" href="?action=edit&id=<%=rs("id")%>&newsid=<%=request("newsid")%>&titles=<%=request("titles")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>&newsid=<%=request("newsid")%>&titles=<%=request("titles")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></div></td>
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
<table width="90%" border=0 align=center cellSpacing=1 class="navi">  
  <tr>
    <td height="20" ><div align="center" class="unnamed1">第<%= page %>页&nbsp; <a href="?page=1&newsid=<%=request("newsid")%>&titles=<%=request("titles")%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh">首页</a> &nbsp;共<%=rs.PageCount%>页&nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %>&newsid=<%=request("newsid")%>&titles=<%=request("titles")%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >上一页</a>
            <% else %>
        上一页
        <% end if %>
&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=page+1%>&newsid=<%=request("newsid")%>&titles=<%=request("titles")%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >下一页</a>
        <% else %>
        下一页
        <% end if %>
&nbsp;<select name="select" onChange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%For m = 1 To rs.PageCount%>
        <option value="?page=<%=m%>&newsid=<%=request("newsid")%>&titles=<%=request("titles")%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" <%if page=m then%>selected<%end if%>><%=m%></option>
        <% Next %>
      </select>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%>&newsid=<%=request("newsid")%>&titles=<%=request("titles")%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >末页</a>
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
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="product_add1.asp?BigClassName=<%=request("BigClassName")%>">新增信息</A></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from product1 where id="& request("id")
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
<script type="text/javascript">
	function check(){
		if(document.formm.title.value == ""){
			alert("请输入标题");
			return false;
		}


		return true;
	}
</script>

<%set rsx=server.createobject("adodb.recordset")
sql = "select * from product_class1 where pid<>0 order by ID asc"
rsx.open sql,conn,1,1
%>

<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rsx.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rsx("classname"))%>",<%= rsx("pid")%>,"<%= trim(rsx("id"))%>");
        <%
        count = count + 1
        rsx.movenext
        loop
        rsx.close
        %>
onecount=<%=count%>;

   function changelocation(locationid)
    {
	var m=0;
    document.formm.SmallClassName.length = 0; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
			 if (m==0)
			{
			m=subcat[i][2];
			}
                document.formm.SmallClassName.options[document.formm.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
		changelocationclass(m);
    }    

</script>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>&BigClassName=<%=request("BigClassName")%>" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="126"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">标题：</div></td>
      <td width="934" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;" value="<%=rs("title")%>"></td>
    </tr>
      <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">类别：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	          <select name="class" id="class" >
          <%set rs1=server.createobject("adodb.recordset")
sql = "select * from product_class1 where pid=0 order by id asc"
rs1.open sql,conn,1,1
dim selclass
selclass=rs("productclass1")
i=0
do while not rs1.eof
i=i+1
%>
          <option value="<%=rs1("id")%>" <%if rs("productclass1")=rs1("id") then%> selected="selected"<%end if%>><%=rs1("classname")%></option>
          <%
rs1.movenext
loop
%>
        </select></td>
      </tr>

    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">说明：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"><%=rs("content")%></textarea>
	                <iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe></td>
    </tr>
	    <input name="image" type="hidden" id="image" value="<%=rs("pic")%>"><tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">上传相关图片<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe>
      (150*110)</td>
    </tr>			
	<input name="image2" type="hidden" id="image2" value="<%=rs("xinghao")%>"/>
		    <tr>
      <td width="126"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">首页显示：</div></td>
      <td width="934" height="20" colspan="2" bgcolor="#FFFFFF"><input name="tj" type="checkbox" id="tj" value="1" <%if rs("tj")=1 then%>checked="checked"<%end if%>/>
        是</td>
	    </tr>
    <tr bgcolor="#A4B6D7">
      <td width="126" height="20" align="right" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <div align="right"></div></td>
      <td height="20" align="left" bgcolor="#FFFFFF">
	  <input type="submit" name="Submit" value="确定" />	  </td>
    </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from product1 where id="&request("id")
		rs1.open sql,conn,1,3
		rs1("title")=request("title")'标题
		rs1("content")=request("theme")'内容
rs1("productclass1")=request("class")'产品类别
		rs1("pic")=request("image")'产品的存放文件名
		rs1("admin")=request("admin")'产品的存放文件
		rs1("xinghao")=request("image2")'产品的存放文件
		rs1("sycx")=request("sycx")'产品的存放文件
		rs1("cd")=request("cd")'产品的存放文件
		rs1("price")=request("price")'产品的存放文件
		rs1("tj")=int(request("tj"))'产品的存放文件
		rs1("time")=now()
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		Response.Write "<script>alert('您已经成功修改');location='product_list1.asp?BigClassName="&request("BigClassName")&"&newsid="&request("newsid")&"&titles="&request("titles")&"'</script>"
end if
if request("action")="del" then
		sql="DELETE  from product1 where id in("& request("id")&")" 
		conn.execute(sql)
response.Redirect "product_list1.asp?page="&request("page")&"&newsid="&request("newsid")&"&titles="&request("titles")
end if
if request("action")="tuij" then
    theme=trim(request("theme"))
    set rs=server.createobject("adodb.recordset")
	sql="select * from product1 where id="&request("id") 
	rs.open sql,conn,1,3
	rs("tj")=0
	rs.update
	rs.close
	set rs=nothing	
	Response.Redirect "product_list1.asp?page="&request("page")
end if
if request("action")="bt" then
    theme=trim(request("theme"))
    set rs=server.createobject("adodb.recordset")
	sql="select * from product1 where id="&request("id") 
	rs.open sql,conn,1,3
	rs("tj")=1
	rs.update
	rs.close
	set rs=nothing	
	Response.Redirect "product_list1.asp?page="&request("page")
end if
if request("action")="sy" then
		set rs=server.createObject("ADODB.Recordset")
		if request("newsid")="" then
		sql="select * from product1 where rootid<"& request("rootid") &" order by rootid desc"
		else
		sql="select * from product1 where rootid<"& request("rootid") &" and productclass1="&request("newsid")&" order by rootid desc"
		end if
		rs.open sql,conn,2,3
		if not rs.eof then
		rootid=rs("rootid")
		set rs1=server.createObject("ADODB.Recordset")
		sql="select * from product1 where rootid="& request("rootid") 
		rs1.open sql,conn,2,3	
		rs("rootid")=rs1("rootid")
		rs1("rootid")=rootid
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		end if
		response.Redirect "product_list1.asp?page="&request("page")&"&newsid="&request("newsid")&"&titles="&request("titles")
end if
if request("action")="xy" then
		set rs=server.createObject("ADODB.Recordset")
		if request("newsid")="" then
		sql="select * from product1 where rootid>"& request("rootid") &" order by rootid asc"
		else
		sql="select * from product1 where rootid>"& request("rootid")&" and productclass1="&request("newsid")&" order by rootid asc"
		end if
		rs.open sql,conn,2,3
		if not rs.eof then
		rootid=rs("rootid")
		set rs1=server.createObject("ADODB.Recordset")
		sql="select * from product1 where rootid="& request("rootid") 
		rs1.open sql,conn,2,3	
		rs("rootid")=rs1("rootid")
		rs1("rootid")=rootid
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		end if
		response.Redirect "product_list1.asp?page="&request("page")&"&newsid="&request("newsid")&"&titles="&request("titles")
end if
%>	
</body>
</html>

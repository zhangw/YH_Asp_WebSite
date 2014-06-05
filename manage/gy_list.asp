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
									信息管理</td>
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
                  <TD height="21"><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle">信息的增加、修改和删除。</TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
                <TR>
                  <TD></TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
			
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th bgcolor="#e0e9f8"  class="STYLE5">类型</th>
    <th height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">产品名称</div></th>
    <th width="146" bgcolor="#e0e9f8"  class="STYLE5">所属类别</th>
    <th width="146" bgcolor="#e0e9f8"  class="STYLE5">联系人</th>
    <th width="146" bgcolor="#e0e9f8"  class="STYLE5">发布时间</th>
    <th width="94" bgcolor="#e0e9f8"  class="STYLE5">状态</th>
    <th width="94" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">操作</div></th>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from gongqiu " 
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
  <tr <%if rowcount mod 2=0 then%>bgcolor="#fffff0"<%else%>bgcolor="#ffffff"<%end if%> class="unnamed1">
    <td  class="STYLE5" style="padding-left:5px;"><%=rs("type")%></td>
    <td height="20"  class="STYLE5" style="padding-left:5px;"><%if len(rs("title"))>20 then%><%=mid(rs("title"),1,20)%>...<%else%><%=rs("title")%><%end if%></td>
	<%
	str=""
	set rs1=server.createobject("adodb.recordset")
	sql="select * from product_class where id="&rs("productclass1")
	rs1.open sql,conn,1,3
	if not rs1.eof then  str=rs1("classname")& " >> "	
	set rs2=server.createobject("adodb.recordset")
	sql="select * from product_class where id="&rs("productclass2")
	rs2.open sql,conn,1,3
	if not rs1.eof then  str=str&rs2("classname")	
    %>
	<td width="146" align="center"  class="STYLE5"><div align="center"><%=str%></div></td>

	<td width="146" align="center"  class="STYLE5"><nobr><%=rs("lxr")%></nobr></td>

	<td width="146" align="center"  class="STYLE5"><%=rs("time")%></td>
    <td width="94" align="center" ><%if rs("sh")=1 then%>
      <a href="?action=tuij&id=<%=rs("id")%>&page=<%=int(request("page"))%>">通过审核</a>
      <%else%>
      <a href="?action=bt&id=<%=rs("id")%>&page=<%=int(request("page"))%>">未通过审核</a>
    <%end if%></td>
    <td width="94" ><div align="center"><A title="修改" href="?action=edit&id=<%=rs("id")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></div></td>
  </tr>
  <%RowCount=RowCount-1
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

            <TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD></TD>
                </TR>
                <TR>
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="product_add.asp">新增信息</A></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from gongqiu where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 修改产品信息</td>
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
			alert("请输入产品的名称");
			return false;
		}
		if(document.formm.admin.value == ""){
			alert("请选择发布人");
			return false;
		}

		return true;
	}
</script>
<%set rsx=server.createobject("adodb.recordset")
sql = "select * from product_class where pid<>0 order by ID asc"
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
<script language=javascript> 
function OnClick(){ 
S=window.showModalDialog("yonghu.asp"); 
document.all.item("admin").value=S; 
} 
</script>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">信息类型：</td>
      <td height="20" colspan="2" bgcolor="#FFFFFF"><select name="type" id="type">
        <option value="供应" <%if rs("type")="供应" then%>selected="selected"<%end if%>>供应</option>
        <option value="求购"  <%if rs("type")="求购" then%>selected="selected"<%end if%>>求购</option>
        <option value="招商"  <%if rs("type")="招商" then%>selected="selected"<%end if%>>招商</option>
        <option value="代理"  <%if rs("type")="代理" then%>selected="selected"<%end if%>>代理</option>
      </select>
      </td>
    </tr>
    <tr>
      <td width="126"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品名称：</div></td>
      <td width="934" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;" value="<%=rs("title")%>"></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品类别：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	          <select name="BigClassName" onChange="changelocation(document.formm.BigClassName.options[document.formm.BigClassName.selectedIndex].value)">
          <%set rs1=server.createobject("adodb.recordset")
sql = "select * from product_class where pid=0 order by id asc"
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
        </select>
       <select name="SmallClassName" >
          <%
		  set rse=server.createobject("adodb.recordset")
			sql="select * from product_class where pid=" & selclass 
			rse.open sql,conn,1,1
			selclass=rse("id")
			if not(rse.eof and rse.bof) then
			%>
          <option value="<%=rse("id")%>"  <%if rs("productclass2")=rse("id") then%> selected="selected"<%end if%>><%=rse("classname")%></option>
          <% rse.movenext
				do while not rse.eof%>
          <option value="<%=rse("id")%>" <%if rs("productclass2")=rse("id") then%> selected="selected"<%end if%>><%=rse("classname")%></option>
          <%
			    	rse.movenext
				loop
			end if
	        rse.close
			%>
        </select></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品型号：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="xinghao" type="text" id="xinghao" style="width:300px;" value="<%=rs("xinghao")%>" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">适用车型：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="sycx" type="text" id="sycx" style="width:300px;"  value="<%=rs("sycx")%>"/></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品产地：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="cd" type="text" id="cd" style="width:300px;"  value="<%=rs("cd")%>"/></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品价格：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="price" type="text" id="price" style="width:300px;"  value="<%=rs("price")%>"/></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品说明：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"><%=rs("theme")%></textarea>
	                <iframe id="editor2" src="../pic/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe></td>
    </tr>
	    <input name="image" type="hidden" id="image" value="<%=rs("pic")%>"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">上传文件图片<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe> </td>
    </tr>			
	   
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3"><div align="right">联系人:</div></td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="lxr" type="text" id="lxr" size="40" value="<%=rs("lxr")%>"/></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">联系电话：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="dh" type="text" id="dh" size="40" value="<%=rs("dh")%>"/></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">联系手机：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="mib" type="text" id="mib" size="40" value="<%=rs("mib")%>"/></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">联系邮箱：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="mail" type="text" id="mail" size="40" value="<%=rs("mail")%>"/></td>
    </tr>
    <tr>
      <td width="126" align="right" bgcolor="#FCFCFC" class="front3">联系地址：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="dizhi" type="text" id="dizhi" size="40" value="<%=rs("dizhi")%>"/></td>
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
		sql="select * from gongqiu where id="&request("id")
		rs1.open sql,conn,1,3
		rs1("title")=request("title")'标题
		rs1("content")=request("theme")'内容
		rs1("productclass1")=request("BigClassName")'产品类别
		rs1("productclass2")=request("SmallClassName")'产品类别
		rs1("pic")=request("image")'产品的存放文件名
		rs1("admin")=request("admin")'产品的存放文件
		rs1("xinghao")=request("xinghao")'产品的存放文件
		rs1("sycx")=request("sycx")'产品的存放文件
		rs1("cd")=request("cd")'产品的存放文件
		rs1("price")=request("price")'产品的存放文件
		rs1("type")=request("type")'产品的存放文件
		rs1("lxr")=request("lxr")'产品的存放文件
		rs1("dh")=request("dh")'产品的存放文件
		rs1("mib")=request("mib")'产品的存放文件
		rs1("mail")=request("mail")'产品的存放文件
		rs1("dizhi")=request("dizhi")'产品的存放文件		
		rs1("time")=date()
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		Response.Write "<script>alert('您已经成功修改');location='gy_list.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from gongqiu where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('删除成功');location='gy_list.asp'</script>"
end if
if request("action")="tuij" then
    theme=trim(request("theme"))
    set rs=server.createobject("adodb.recordset")
	sql="select * from gongqiu where id="&request("id") 
	rs.open sql,conn,1,3
	rs("sh")=0
	rs.update
	rs.close
	set rs=nothing	
	Response.Redirect "gy_list.asp?page="&request("page")
end if
if request("action")="bt" then
    theme=trim(request("theme"))
    set rs=server.createobject("adodb.recordset")
	sql="select * from gongqiu where id="&request("id") 
	rs.open sql,conn,1,3
	rs("sh")=1
	rs.update
	rs.close
	set rs=nothing	
	Response.Redirect "gy_list.asp?page="&request("page")
end if
%>		
</body>
</html>

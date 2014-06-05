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
									招聘管理</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>

<%
if request("action")="list" then
set rs=server.createobject("adodb.recordset")
sql="select * from job order by id desc"
rs.open sql,conn,1,1
%>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td  align="center" valign="top">

         <%
	if rs.eof then 
	response.Write "没有数据"
	else	 
		  do until rs.eof %>
		 <table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#c0c0c0" class="border" style="word-break:break-all;margin-top:10px;">
        <tr bgcolor="#A4B6D7" class="title"> 
          <td width="70" height="25" align="center" bgcolor="#e0e9f8">职位名称：</td>
          <td width="165"  height="25" align="left" bgcolor="#ECF5FF">&nbsp;<%=rs("jobname")%></td>
          <td width="86" align="center" bgcolor="#e0e9f8" >工作地点：</td>
          <td width="235" align="left" bgcolor="#ECF5FF" >&nbsp;<%=rs("adress")%></td>
          <td width="87" align="center" bgcolor="#e0e9f8" >招聘人数：</td>
          <td width="227" align="left" bgcolor="#ECF5FF" >&nbsp;<%=rs("shu")%></td>
          </tr>
        <tr class="tdbg">
          <td height="106" align="center" bgcolor="#e0e9f8">职位要求：</td>
          <td colspan="5" align="left" bgcolor="#ECF5FF" style="padding-left:5px;"><%=rs("content")%></td>
          </tr>
        <tr class="tdbg"> 
          <td width="70" height="22" align="center" bgcolor="#e0e9f8">发布时间：</td>
          <td width="165" align="left" bgcolor="#ECF5FF">&nbsp;<%=rs("shijian")%></td>
          <td colspan="3" align="left"  bgcolor="#ECF5FF">&nbsp;&nbsp;</td>
          <td align="right"  bgcolor="#ECF5FF">操作：&nbsp;<a href="?action=edit&id=<%=rs("id")%>">修改</a> <a href="?action=del&id=<%=rs("id")%>">删除</a>&nbsp;&nbsp;</td>
        </tr>
	  </table>
  <% rs.movenext
  loop
  end if
  %>
      
      
    </td>
  </tr>
</table>
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
end if
%>
						<script language=javascript>
function formCheck()
{	
	if (document.form1.jobname.value == "")
	{
		alert("请填写职位名称")
		document.form1.jobname.focus()
		return false
	}
	if (document.form1.adress.value == "")
	{
		alert("请填写工作地点")
		document.form1.adress.focus()
		return false
	}
	if (document.form1.shu.value == "")
	{
		alert("请填写招聘人数")
		document.form1.shu.focus()
		return false
	}
		if (document.form1.content.value == "")
	{
		alert("请填写职位要求")
		document.form1.content.focus()
		return false
	}
		if (document.form1.Content.value == "")
	{
		alert("请填写留言内容")
		document.form1.Content.focus()
		return false
	}
}
</script>
<%if request("action")="add" then%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td  align="center" valign="top">


		 <table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#c0c0c0" class="border" style="word-break:break-all;margin-top:10px;">
		   <form id="form1" name="form1" method="post" action="?action=addsave" onSubmit="return formCheck(); ">
        <tr bgcolor="#A4B6D7" class="title"> 
          <td width="92" height="22" align="center" bgcolor="#e0e9f8">职位名称：</td>
          <td width="160"  height="20" align="left" bgcolor="#ECF5FF">&nbsp;
            
            <input type="text" name="jobname" style="height:16px;"/>          </td>
          <td width="72" height="20" align="center" bgcolor="#e0e9f8" >工作地点：</td>
          <td width="224" height="20" align="left" bgcolor="#ECF5FF" >&nbsp;
            <input type="text" name="adress" style="height:16px;"/></td>
          <td width="96" height="20" align="center" bgcolor="#ECF5FF" >招聘人数：</td>
          <td width="265" height="20" align="left" bgcolor="#ECF5FF" >&nbsp;
            <input type="text" name="shu" style="height:16px;"/></td>
          </tr>
        <tr class="tdbg">
          <td height="101" align="center" bgcolor="#e0e9f8">职位要求：</td>
          <td height="101" colspan="5" align="left" bgcolor="#ECF5FF">&nbsp;
            <textarea name="content" cols="100" rows="3" style="height:90px;"></textarea>
            &nbsp;</td>
          </tr>
        <tr class="tdbg"> 
          <td width="92" height="20" align="center" bgcolor="#e0e9f8">发布时间：</td>
          <td height="20" align="left" bgcolor="#ECF5FF">&nbsp;
            <input type="text" name="shijian" value="<%=date()%>" style="height:16px;"/></td>
          <td height="20" colspan="4" align="left"  bgcolor="#ECF5FF">&nbsp;&nbsp;
            <input type="submit" name="Submit" value="添加" /></td>
          </tr> </form>
	  </table>

      
      
    </td>
  </tr>
</table>
<%end if
if request("action")="addsave" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from job" 
	rs.open sql,conn,1,3
	rs.addnew
	rs("jobname")=trim(request("jobname"))
	rs("adress")=request("adress")
	rs("shu")=request("shu")
	rs("content")=request("content")
	rs("shijian")=request("shijian")
	rs("ren")=0
	rs.update
	rs.close
	set rs=nothing	
	Response.Write "<script>alert('添加成功');location='job.Asp?action=add'</script>"
end if
%>
<%if request("action")="edit" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from job where id="&request("id") 
	rs.open sql,conn,1,3
%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td  align="center" valign="top"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#c0c0c0" class="border" style="word-break:break-all;margin-top:10px;">
		   <form id="form1" name="form1" method="post" action="?action=editsave&id=<%=request("id")%>" onSubmit="return formCheck(); ">
        <tr bgcolor="#A4B6D7" class="title"> 
          <td width="109" height="22" align="center" bgcolor="#e0e9f8">职位名称：</td>
          <td width="172"  height="25" align="left" bgcolor="#ECF5FF">&nbsp;
            
              <input type="text" name="jobname" value="<%=rs("jobname")%>" style="height:16px;"/>
           
          </td>
          <td width="104" align="center" bgcolor="#e0e9f8" >工作地点：</td>
          <td width="286" align="left" bgcolor="#ECF5FF" >&nbsp;
            <input type="text" name="adress" value="<%=rs("adress")%>" style="height:16px;"/></td>
          <td width="105" align="center" bgcolor="#e0e9f8" >招聘人数：</td>
          <td width="284" align="left" bgcolor="#ECF5FF" >&nbsp;
            <input type="text" name="shu" value="<%=rs("shu")%>" style="height:16px;"/></td>
          </tr>
        <tr class="tdbg">
          <td height="22" align="center" bgcolor="#e0e9f8">职位要求：</td>
          <td height="60" colspan="5" align="left" bgcolor="#ECF5FF">&nbsp;
            <textarea name="content" cols="100" rows="3" style="height:50px;"><%=rs("content")%></textarea>
            &nbsp;</td>
          </tr>
        <tr class="tdbg"> 
          <td width="109" height="22" align="center" bgcolor="#e0e9f8">发布时间：</td>
          <td width="172" align="left" bgcolor="#ECF5FF">&nbsp;
            <input type="text" name="shijian" value="<%=rs("shijian")%>" style="height:16px;"/>            &nbsp;</td>
          <td colspan="3" align="left"  bgcolor="#ECF5FF">&nbsp;&nbsp;
            <input type="submit" name="Submit" value="修改" /></td>
          <td align="right"  bgcolor="#ECF5FF">&nbsp;</td>
        </tr> </form>
	  </table>

      
      
    </td>
  </tr>
</table>
<%end if
if request("action")="editsave" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from job where id="&request("id") 
	rs.open sql,conn,1,3
	rs("jobname")=trim(request("jobname"))
	rs("adress")=request("adress")
	rs("shu")=request("shu")
	rs("content")=request("content")
	rs("shijian")=request("shijian")
	rs.update
	rs.close
	set rs=nothing	
	Response.Write "<script>alert('修改成功');location='job.Asp?action=list'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from job where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		response.Redirect "job.asp?action=list"
end if
%>


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
									应聘管理</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>

<%
if request("action")="list" or request("action")="" then
set rs=server.createobject("adodb.recordset")
sql="select * from jl order by id desc"
rs.open sql,conn,1,1
%>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td  align="center" valign="top"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#c0c0c0" class="border" style="word-break:break-all;margin-top:10px;">   
      <tr bgcolor="#A4B6D7" class="title"> 
          <td width="106" height="25" align="center" bgcolor="#e0e9f8"><strong>姓名</strong></td>
          <td width="100" align="center" bgcolor="#e0e9f8"><strong>性别</strong></td>
          <td width="100"  height="25" align="center" bgcolor="#e0e9f8"><strong>&nbsp;学历</strong></td>
          <td width="150" align="center" bgcolor="#e0e9f8" ><strong>专业</strong></td>
          <td width="150" align="center" bgcolor="#e0e9f8" ><strong>联系电话</strong></td>
          <td width="150" align="center" bgcolor="#e0e9f8" ><strong>&nbsp;面试职位</strong></td>
          <td width="150" align="center" bgcolor="#e0e9f8" ><strong>操作</strong></td>
        </tr><%
	if rs.eof then 
	response.Write "没有数据"
	else	
	i=0 
		  do until rs.eof %>
		
       
        <tr class="tdbg" bgcolor="#FFFFFF">
     
          <td width="106" align="center" bgcolor="#FFFFFF" <%if i mod 2=0 then%>bgcolor="#fffff0"<%end if%> ><%=rs("TrueName")%></td>
          <td width="100" align="center" bgcolor="#FFFFFF" <%if i mod 2=0 then%>bgcolor="#fffff0"<%end if%>><%=rs("Sex")%></td>
          <td width="100" align="center" <%if i mod 2=0 then%>bgcolor="#fffff0"<%end if%> ><%=rs("StudyStory")%></td>
          <td width="150" align="center" bgcolor="#FFFFFF" <%if i mod 2=0 then%>bgcolor="#fffff0"<%end if%>><%=rs("Speciality")%></td>
          <td width="150" align="center" bgcolor="#FFFFFF" <%if i mod 2=0 then%>bgcolor="#fffff0"<%end if%>><%=rs("Phone")%></td>
          <td width="150" align="center" bgcolor="#FFFFFF" <%if i mod 2=0 then%>bgcolor="#fffff0"<%end if%>><%=rs("zhiwei")%></td>
		       <td width="150" height="22" align="center" bgcolor="#FFFFFF" <%if i mod 2=0 then%>bgcolor="#fffff0"<%end if%>><a href="?action=edit&id=<%=rs("id")%>">查看</a>&nbsp;<a href="?action=del&id=<%=rs("id")%>">删除</a></td>
        </tr>
		  
  <% rs.movenext
  loop
  end if
  %></table>
      
      
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
<%if request("action")="edit" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from jl where id="&request("id") 
	rs.open sql,conn,1,3
%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td  align="center" valign="top"><table width="800" border="0" cellspacing="0" cellpadding="0">
                <form name="form1" method="post" action="?"> <tr>
                  <td><TABLE borderColor="#111111" cellSpacing="0" cellPadding="0" width="96%" border="0">
                    <TBODY>
                      <TR>
                        <TD height="20"></TD>
                        <TD>&nbsp;</TD>
                        <TD align="right"><a href="?action=del&id=<%=request("id")%>">删除</a></TD>
                      </TR>
                      <TR>
                        <TD width="40%" height="20">
                        </TD>
                        <TD width="43%" class="muen">职位名称：<%=rs("zhiwei")%></TD>
                        <TD align="right" width="17%">&nbsp;</TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                    <TABLE borderColor="#111111" cellSpacing="0" cellPadding="0" width="100%" border="0">
                      <TBODY>
                        <TR>
                          <TD width="100%">&nbsp;</TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <TABLE width="96%" border="0" cellPadding="0" cellSpacing="1" bgcolor="#c0c0c0">
                      <TBODY>
                        <TR>
                          <TD height="25" colSpan="4" bgColor="#e0e9f8" class="main"><strong>基本信息</strong></TD>
                        </TR>
                        <TR>
                          <TD width="20%" height="25" bgColor="#F3F7FC" class="main"> 姓名*</TD>
                          <TD width="32%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("TrueName")%></TD>
                          <TD width="15%" height="25" bgColor="#F3F7FC" class="main"> 性别</TD>
                          <TD width="33%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("Sex")%></TD>
                        </TR>
                        <TR>
                          <TD width="20%" height="25" bgColor="#F3F7FC" class="main"> 学历*</TD>
                          <TD width="32%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("StudyStory")%></TD>
                          <TD width="15%" height="25" bgColor="#F3F7FC" class="main"> 专业*</TD>
                          <TD width="33%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("Speciality")%></TD>
                        </TR>
                        <TR>
                          <TD width="20%" height="25" bgColor="#F3F7FC" class="main"> 户口所在地*</TD>
                          <TD width="32%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("RPR")%>
                          </TD>
                          <TD width="15%" height="25" bgColor="#F3F7FC" class="main"> 身高(cm)</TD>
                          <TD width="33%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("Stature")%></TD>
                        </TR>
                        <TR>
                          <TD width="20%" height="25" bgColor="#F3F7FC" class="main"> 出生日期*</TD>
                          <TD height="25" colSpan="3" bgcolor="#FFFFFF" class="main"><%=rs("Birthday")%></TD>
                        </TR>
                        <TR>
                          <TD width="20%" height="25" bgColor="#F3F7FC" class="main"> 证件类型</TD>
                          <TD width="32%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("CardType")%></TD>
                          <TD width="15%" height="25" bordercolor="#F3F7FC" bgColor="#F3F7FC" class="main"> 证件号</TD>
                          <TD width="33%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("CardNo")%></TD>
                        </TR>
                        <TR>
                          <TD width="20%" height="25" bgColor="#F3F7FC" class="main"> 婚姻状况</TD>
                          <TD width="32%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("Marriage")%></TD>
                          <TD width="15%" height="25" bordercolor="#F3F7FC" bgColor="#F3F7FC" class="main"> 健康状况</TD>
                          <TD width="33%" height="25" bgcolor="#FFFFFF" class="main"><%=rs("infornametion")%></TD>
                        </TR>
                        <TR>
                          <TD height="25" bgColor="#F3F7FC" class="main"> 目前收入(元/年)</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("Income")%></TD>
                          <TD height="25" bordercolor="#F3F7FC" bgColor="#F3F7FC" class="main"> 工作年限</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("JobYear")%></TD>
                        </TR>
                        <TR>
                          <TD height="25" bgColor="#F3F7FC" class="main"> 居住地(省)</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("Province")%></TD>
                          <TD height="25" bordercolor="#F3F7FC" bgColor="#F3F7FC" class="main"> 居住地(市)</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("City")%></TD>
                        </TR>
                        <TR>
                          <TD height="25" bgColor="#F3F7FC" class="main"> 电子邮箱</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("Email")%></TD>
                          <TD height="25" bordercolor="#F3F7FC" bgColor="#F3F7FC" class="main">联系电话*</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("Phone")%></TD>
                        </TR>
                        <TR>
                          <TD height="25" bgColor="#F3F7FC" class="main"> 通讯地址</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("Address")%></TD>
                          <TD height="25" bordercolor="#F3F7FC" bgColor="#F3F7FC" class="main"> 邮编</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("Postalcode")%></TD>
                        </TR>
                        <TR>
                          <TD height="25" bgColor="#F3F7FC" class="main"> 职称</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("Duty")%></TD>
                          <TD height="25" bordercolor="#F3F7FC" bgColor="#F3F7FC" class="main"> 期望薪水</TD>
                          <TD height="25" bgcolor="#FFFFFF" class="main"><%=rs("ExpPay")%></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <br />
                    <TABLE width="96%" border="0" cellPadding="0" cellSpacing="1" borderColor="#cccccc" bgcolor="#c0c0c0">
                      <TBODY>
                        <TR>
                          <TD width="100%" height="25" bgColor="#E0E9F8" class="main"><strong>学业情况</strong></TD>
                        </TR>
                        <TR>
                          <TD height="25" align="left" bgcolor="#FFFFFF" class="main"><%=replace(rs("Study"),chr(13)&chr(10),"<br>")%></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <br />
                    <TABLE width="96%" border="0" cellPadding="0" cellSpacing="1" borderColor="#cccccc" bgcolor="#c0c0c0">
                      <TBODY>
                        <TR>
                          <TD width="100%" height="25" bgColor="#E0E9F8"><p class="main"><strong>工作经历</strong></p></TD>
                        </TR>
                        <TR>
                          <TD height="25" align="left" bgcolor="#FFFFFF"><%=replace(rs("JobStory"),chr(13)&chr(10),"<br>")%></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <br />
                    <TABLE width="96%" border="0" cellPadding="0" cellSpacing="1" borderColor="#cccccc" bgcolor="#c0c0c0">
                      <TBODY>
                        <TR>
                          <TD width="100%" height="25" bgColor="#E0E9F8"><p class="main"><strong>其它介绍</strong></p></TD>
                        </TR>
                        <TR>
                          <TD height="25" align="left" bgcolor="#FFFFFF"><%=replace(rs("Other"),chr(13)&chr(10),"<br>")%></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                    <TABLE borderColor="#111111" cellSpacing="0" cellPadding="0" width="100%" border="0">
                      <TBODY>
                        <TR>
                          <TD vAlign="bottom" width="100%" height="30"><p align="center">
                            <INPUT type="submit" value="返回" name="Submit2">
                          <a href="?action=del&id=<%=request("id")%>">删除</a></p></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                </tr>
				</form>
      </table>

      
      
    </td>
  </tr>
</table>
<%end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from jl where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		response.Redirect "yp.asp?action=list"
end if
%>


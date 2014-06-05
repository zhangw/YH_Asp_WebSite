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
       Sub MainFl() 
           Dim Rs 
           Set Rs = Conn.Execute("SELECT helpid,title FROM helpid") 
           If Not Rs.Eof Then 
               Do While Not Rs.Eof 
                   response.Write "<option value=""" & Trim(Rs("helpid")) & """>" & Trim(Rs("title")) & "</option>" 
                   response.Write "<br/>" 
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
									产品管理&gt;&gt;发布产品</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>
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

<form name="formm" method="post" action="?action=aa" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="100"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品名称：</div></td>
      <td width="596" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;"></td>
    </tr>
    <tr>
      <td width="100" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品类别：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="BigClassName"><% MainFl()%></select></td>
    </tr>
    <tr>
      <td width="100" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品说明：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"></textarea>
	                <iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe></td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
      <th height="20" align="left" bgcolor="#FFFFFF">
	  <input type="submit" name="Submit" value="确定" />
	  </th>
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%
if request("action")="aa" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from help "
		rs1.open sql,conn,1,3
		rs1.addnew
		rs1("title")=request("title")'标题
		rs1("content")=request("theme")'内容
		rs1("helpid")=request("BigClassName")'产品类别
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
	Response.Write "<script>alert('您已经成功添加');location='help_add.asp'</script>"
end if	
%>
</body>
</html>


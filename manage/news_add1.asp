<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title></title>
<script language=javascript>
function formCheck()
{	
	if (document.formm.title.value == "")
	{
		alert("请填写标题")
		document.formm.title.focus()
		return false
	}

}
</script>
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
									新闻中心 &gt;&gt; 加入资讯&nbsp;&nbsp;&nbsp;</td>
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
<form name="formm" method="post" action="?action=aa" onSubmit="return formCheck();">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="150"  height="20" bgcolor="#FFFFFF" class="STYLE5"><div align="right">标&nbsp;&nbsp;&nbsp;&nbsp;题：</div></td>
      <td width="596" height="20" colspan="2" align="left" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:550px;"></td>
    </tr>
    <tr>
      <td width="150" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">关 键 字：</td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><input name="gjz" type="text" id="gjz" style="width:550px;" /></td>
    </tr>
		    <tr>
      <td width="150" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">描&nbsp;&nbsp;&nbsp;&nbsp;述：</td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><textarea name="ms" rows="3" id="ms" style="width:550px;"></textarea></td>
		    </tr>
    <tr>
      <td width="150" height="20" align="center" valign="middle" bgcolor="#FFFFFF" class="STYLE5"><div align="right">发 布 人：</div></td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><input name="fabu" type="text" id="fabu" style="width:550px;" value="管理员" /></td>
    </tr>
    <tr>
      <td width="150" height="20" valign="middle" bgcolor="#FFFFFF" class="STYLE5"><div align="right">发布时间：</div></td>
      <td colspan="2" align="left" bgcolor="#FFFFFF"><input name="time" type="text" id="time" value="<%=now()%>" style="width:550px;"/></td>
    </tr>
    <tr>
      <td width="150" height="20" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">内&nbsp;&nbsp;&nbsp;&nbsp;容：</div></td>
      <td colspan="2" align="left" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"></textarea>
      <iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="560" height="405"></iframe></td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <th width="150" height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <div align="right"></div></th>
      <th height="20" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="确定添加" /></th>
      <th height="20" align="left" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%
if request("action")="aa" then
    pic=trim(request("image"))
    theme=trim(request("theme"))
    set rs=server.createobject("adodb.recordset")
	sql="select * from news" 
	rs.open sql,conn,1,3
	rs.addnew
	rs("title")=trim(request("title"))
	rs("content")=theme
	rs("newsid")=2
	rs("time")=request("time")
	rs("pic")=request("image")
    rs("zxdt")=int(request("zxdt"))
	rs("gjz")=request("gjz")
	rs("ms")=request("ms")
	rs("tthg")=int(request("tthg"))
	rs("tpxw")=int(request("tpxw"))
	rs("fabu")=request("fabu")
	rs("yd")=0
	rs.update
	rs.close
	set rs=nothing
	

	Response.Write "<script>alert('您已经成功添加');location='news_add1.asp'</script>"
end if	
%>
</body>
</html>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>乔恩传媒</title>
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
									会员管理 &gt;&gt; 添加会员&nbsp;&nbsp;&nbsp;</td>
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
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
      <form id="formm" name="formm" method="post" action="?action=aa&id=<%=request("id")%>&page=<%=request("page")%>">
      <tr>
        <td width="150" height="30" align="right" class="css"><span class="css3">*</span>公司名称：&nbsp;&nbsp;</td>
        <td width="746" align="left" class="css3"><input name="company" type="text" id="company" size="40" style="height:18px;" value="<%=rs("company")%>"/>
          &nbsp;此项为必填项，注册企业请填写工商局注册的全称。</td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css">企业类型：&nbsp;&nbsp;</td>
        <td align="left" class="css"><select name="class" id="class" style="height:18px;">
          <option value="4S店-乘用车辆" <%if rs("class")="4S店-乘用车辆" then%>selected="selected"<%end if%>>4S店-乘用车辆</option>
          <option value="4S店-商务车辆" <%if rs("class")="4S店-商务车辆" then%>selected="selected"<%end if%>>4S店-商务车辆</option>
          <option value="4S店-工程机械车辆" <%if rs("class")="4S店-工程机械车辆" then%>selected="selected"<%end if%>>4S店-工程机械车辆</option>
          <option value="快修服务" <%if rs("class")="快修服务" then%>selected="selected"<%end if%>>快修服务</option>
          <option value="上牌服务" <%if rs("class")="上牌服务" then%>selected="selected"<%end if%>>上牌服务</option>
          <option value="车辆检测" <%if rs("class")="车辆检测" then%>selected="selected"<%end if%>>车辆检测</option>
        </select>        </td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css">详细地址：&nbsp;&nbsp;</td>
        <td align="left" class="css3"><input name="adress" type="text" id="adress" size="40" style="height:18px;" value="<%=rs("adress")%>"/>
          &nbsp;此项为必填项。</td>
        </tr>

      <tr>
        <td width="150" height="30" align="right" class="css"><span class="css3">*</span>会员登录名：&nbsp;&nbsp;</td>
        <td align="left" class="css3"><input name="admin" type="text" id="admin" size="20" style="height:18px;" value="<%=rs("admin")%>"/>
          &nbsp;一旦注册成功不可以修改。</td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css"><span class="css3">*</span>密码：&nbsp;&nbsp;</td>
        <td align="left" class="css3"><input name="pw1" type="password" id="pw1" size="21" style="height:18px;" value="<%=rs("password")%>"/>
          &nbsp;建议采用易记、难猜的组合。</td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css"><span class="css3">*</span>重新输入密码：&nbsp;&nbsp;</td>
        <td align="left" class="css3"><input name="pw2" type="password" id="pw2" size="21" style="height:18px;" value="<%=rs("password")%>"/>
          &nbsp;请再输入一次您上面填写的密码。</td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css"><span class="css3">*</span>真实姓名：&nbsp;&nbsp;</td>
        <td align="left" class="css"><input name="name" type="text" id="name" style="height:18px;" value="<%=rs("xm")%>"/></td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css">性别：&nbsp;&nbsp;</td>
        <td align="left" class="css"><input name="sex" type="radio" value="先生" checked="checked" />
          先生
            <input type="radio" name="sex" value="女士" />
            女士</td>
      </tr>
      <tr>
        <td width="150" height="30" align="right" class="css">职务：&nbsp;&nbsp;</td>
        <td align="left" class="css"><input name="zhiwu" type="text" id="zhiwu" style="height:18px;" value="<%=rs("zhiwu")%>"/></td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css"><span class="css3">*</span>电子邮件：&nbsp;&nbsp;</td>
        <td align="left" class="css3"><input name="mail" type="text" id="mail" style="height:18px;" value="<%=rs("mail")%>"/>
          &nbsp;很重要，这是客户与您联系最常用的方式，请务必填写正确、常用的邮箱地址。</td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css">固定电话：&nbsp;&nbsp;</td>
        <td align="left" class="css">区号
          <input name="qh1" type="text" id="qh1" size="5" style="height:18px;" value="<%=rs("qh1")%>"/>
          电话号码
          <input name="dh1" type="text" id="dh1" style="height:18px;" value="<%=rs("dh1")%>"/></td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css">传真：&nbsp;&nbsp;</td>
        <td align="left" class="css">区号
          <input name="qh2" type="text" id="qh2" size="5" style="height:18px;" value="<%=rs("qh2")%>"/>
电话号码
<input name="dh2" type="text" id="dh2" value="<%=rs("dh2")%>"/></td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css">手机：&nbsp;&nbsp;</td>
        <td align="left" class="css"><input name="mib" type="text" id="mib" style="height:18px;" value="<%=rs("mib")%>"/></td>
        </tr>
      <tr>
        <td width="150" height="30" align="right" class="css5">服务项目：&nbsp;&nbsp;</td>
        <td height="25" align="left" class="css5"><span class="css">
          <input name="xiangmu" type="text" id="xiangmu" style="height:18px;" value="<%=rs("xiangmu")%>" size="50"/>
        </span></td>
      </tr>
      <tr>
        <td width="150" height="30" align="right" class="css5">公司介绍：&nbsp;&nbsp;</td>
        <td height="25" align="left" class="css5">
          <textarea name="content"  id="content" style="display:none"><%=rs("content")%></textarea>
		  <iframe id="editor2" src="../pic/eWebEditor.asp?id=content" frameborder=1 scrolling=no width="550" height="405"></iframe>      </td>
      </tr>
      <tr>
        <td height="30" align="center">公司标志：</td>
        <td height="30" align="left"><input name="image" type="hidden" id="image"/><iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
      </tr>
      <tr>
        <td width="150" height="30" align="center">&nbsp;</td>
        <td height="30" align="left"><input type="submit" name="Submit" value="添加" /></td>
      </tr>  </form>
    </table>
<%
if request("action")="aa" then
if request("company") = "" then
response.Write "<script>alert('请输入公司名称...');history.back();</script>"
response.End()
end if
if request("admin") = "" then
response.Write "<script>alert('请输入会员登录名...');history.back();</script>"
response.End()
end if
if request("pw1") = "" then
response.Write "<script>alert('请输入密码...');history.back();</script>"
response.End()
end if
if request("pw2") = "" then
response.Write "<script>alert('请再输入一次密码...');history.back();</script>"
response.End()
end if
if request("pw2") <> request("pw1") then
response.Write "<script>alert('两次输入的密码不一样...');history.back();</script>"
response.End()
end if
if request("name") = "" then
response.Write "<script>alert('请输入姓名...');history.back();</script>"
response.End()
end if
if request("mail") = "" then
response.Write "<script>alert('请输入电子邮件...');history.back();</script>"
response.End()
end if
if request("mail") = "" then
response.Write "<script>alert('请输入电子邮件...');history.back();</script>"
response.End()
end if
if Instr(request("mail"),"@") = 0 then
response.Write "<script>alert('请输入正确电子邮件...');history.back();</script>"
response.End()
end if
set rs = server.CreateObject("adodb.recordset")
sql = "select  * from [user] where admin='"&request("admin")&"'order by id asc"
rs.open sql,conn,1,3
if not rs.eof then
response.Write "<script>alert('用户名已经存在，请重新输入...');history.back();</script>"
response.End()
end if
set rs = server.CreateObject("adodb.recordset")
sql = "select  * from [user] order by id asc"
rs.open sql,conn,1,3
rs.addnew
rs("company") = request("company")
rs("class") = request("class")
rs("adress") = request("adress")
rs("admin") = request("admin")
rs("password") = request("pw1")
rs("xm") = request("name")
rs("sex") = request("sex")
rs("zhiwu") = request("zhiwu")
rs("mail") = request("mail")
rs("qh1") = request("qh1")
rs("dh1") = request("dh1")
rs("qh2") = request("qh2")
rs("dh2") = request("dh2")
rs("mib") = request("mib")
rs("pic") = request("image")
rs("time")=now()
rs("ztime")=now()
rs("sh")=1
rs("content") = request("content")
rs("xiangmu") = request("xiangmu")
rs.update
rs.close
set rs = nothing
response.Redirect "user_list.asp"
end if	
%>
</body>
</html>

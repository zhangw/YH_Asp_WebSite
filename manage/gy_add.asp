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
									信息管理&gt;&gt;发布信息</td>
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
<script type="text/javascript">
	function check(){
		if(document.formm.title.value == ""){
			alert("请输入产品的名称");
			return false;
		}
		if(document.formm.lxr.value == ""){
			alert("请输入联系人");
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
<form name="formm" method="post" action="?action=aa" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">信息类型：</td>
      <td height="20" colspan="2" bgcolor="#FFFFFF"><select name="type" id="type">
        <option value="供应" selected="selected">供应</option>
        <option value="求购">求购</option>
        <option value="招商">招商</option>
        <option value="代理">代理</option>
      </select>      </td>
    </tr>
    <tr>
      <td width="126"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品名称：</div></td>
      <td width="934" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;"></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品类别：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	          <select name="BigClassName" onChange="changelocation(document.formm.BigClassName.options[document.formm.BigClassName.selectedIndex].value)">
          <%set rs1=server.createobject("adodb.recordset")
sql = "select * from product_class where pid=0 order by id asc"
rs1.open sql,conn,1,1
dim selclass
selclass=rs1("id")
i=0
do while not rs1.eof
i=i+1
%>
          <option value="<%=rs1("id")%>" ><%=rs1("classname")%></option>
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
          <option value="<%=rse("id")%>" ><%=rse("classname")%></option>
          <% rse.movenext
				do while not rse.eof%>
          <option value="<%=rse("id")%>" ><%=rse("classname")%></option>
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
      <td colspan="2" bgcolor="#FFFFFF"><input name="xinghao" type="text" id="xinghao" style="width:300px;" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">适用车型：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="sycx" type="text" id="sycx" style="width:300px;" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品产地：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="cd" type="text" id="cd" style="width:300px;" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品价格：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="price" type="text" id="price" style="width:300px;" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">产品说明：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"></textarea>
	                <iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe></td>
    </tr>
	    <input name="image" type="hidden" id="image"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">上传文件图片<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe> </td>
    </tr>			
	<input name="image1" type="hidden" id="image1"/>		   
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3"><div align="right">联系人:</div></td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="lxr" type="text" id="lxr" size="40"/></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">联系电话：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="dh" type="text" id="dh" size="40"/></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">联系手机：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="mib" type="text" id="mib" size="40"/></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">联系邮箱：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="mail" type="text" id="mail" size="40"/></td>
    </tr>
    <tr>
      <td width="126" align="right" bgcolor="#FCFCFC" class="front3">联系地址：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="dizhi" type="text" id="dizhi" size="40"/></td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td width="126" height="20" align="right" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <div align="right"></div></td>
      <td height="20" align="left" bgcolor="#FFFFFF">
	  <input type="submit" name="Submit" value="确定" />	  </td>
    </tr>
  </table>
</form>
<%
if request("action")="aa" then

		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from gongqiu "
		rs1.open sql,conn,1,3
		rs1.addnew
		rs1("title")=request("title")'标题
		rs1("type")=request("type")'标题
		rs1("content")=request("theme")'内容
		rs1("productclass1")=request("BigClassName")'产品类别
		rs1("productclass2")=request("SmallClassName")'产品类别
		rs1("pic")=request("image")'产品的存放文件名
		rs1("admin")=request("admin")'产品的存放文件
		rs1("xinghao")=request("xinghao")'产品的存放文件
		rs1("sycx")=request("sycx")'产品的存放文件
		rs1("cd")=request("cd")'产品的存放文件
		rs1("lxr")=request("lxr")'产品的存放文件
		rs1("dh")=request("dh")'产品的存放文件
		rs1("mib")=request("mib")'产品的存放文件
		rs1("mail")=request("mail")'产品的存放文件
		rs1("dizhi")=request("dizhi")'产品的存放文件
		rs1("price")=request("price")'产品的存放文件
		rs1("sh")=1'产品的存放文件
		rs1("time")=date()
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
	Response.Write "<script>alert('您已经成功添加');location='gy_add.asp'</script>"
end if	
%>
</body>
</html>


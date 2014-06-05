<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title></title>
<!--#include file="include/db_conn.asp"-->
<!--#include file="test.asp"-->
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
								<td align="right">您的位置：<A href="index.asp" target="_top">后台管理</A> &nbsp;&nbsp; 
									</td>
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
			alert("请输入标题");
			return false;
		}


		return true;
	}
</script>

<form name="formm" method="post" action="?action=aa" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="126"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">标题：</div></td>
      <td width="934" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;">
      <input type="hidden" name="BigClassName" value="<%=request("BigClassName")%>"/></td>
    </tr>

 <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">类别：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	          <select name="class" id="class" >
          <%set rs1=server.createobject("adodb.recordset")
sql = "select * from product_class where pid=0 order by id asc"
rs1.open sql,conn,1,1
do while not rs1.eof
%>
          <option value="<%=rs1("id")%>" ><%=rs1("classname")%></option>
          <%
rs1.movenext
loop
%>
        </select>      </td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">说明：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"></textarea>
	                <iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe></td>
    </tr>
	    <input name="image" type="hidden" id="image"/>

    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">请上传小图<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe> 
      (150*110)</td>
    </tr>			
	<input name="image1" type="hidden" id="image1"/>			
	<input name="image2" type="hidden" id="image2"/>
		    <tr>
      <td width="126"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">首页显示：</div></td>
      <td width="934" height="20" colspan="2" bgcolor="#FFFFFF"><input name="tj" type="checkbox" id="tj" value="1" />
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
<%
if request("action")="aa" then
max=1
		set rs2=server.CreateObject("ADODB.Recordset")
		sql="select rootid from product order by rootid desc"
		rs2.open sql,conn,1,3
        if int(rs2("rootid"))=0  then
		max=0
		else
		
		max=int(rs2("rootid"))
		end if
		max=max+1
		
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from product "
		rs1.open sql,conn,1,3
		rs1.addnew
		rs1("title")=request("title")'标题
		rs1("rootid")=max'标题
		rs1("content")=request("theme")'内容
		rs1("productclass1")=request("class")'产品类别
		rs1("productclass2")=request("SmallClassName")'产品类别
		rs1("pic")=request("image")'产品的存放文件名
		rs1("admin")=request("admin")'产品的存放文件
		rs1("xinghao")=request("image2")'内容
		rs1("yd")=0
		rs1("sycx")=request("sycx")'产品的存放文件
		rs1("cd")=request("cd")'产品的存放文件
		rs1("price")=request("price")'产品的存放文件
		rs1("tj")=int(request("tj"))'产品的存放文件
		rs1("time")=now()
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
	Response.Write "<script>alert('您已经成功添加');location='product_add.asp?BigClassName="&request("BigClassName")&"'</script>"
end if	
%>
</body>
</html>


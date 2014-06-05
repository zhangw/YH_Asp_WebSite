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
									建材超市&gt;&gt;发布产品</td>
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
<%
		set rs2=server.CreateObject("ADODB.Recordset")
		sql="select * from guanggao where id=1"
		rs2.open sql,conn,1,1
%>
<form name="formm" method="post" action="?action=aa" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
	    <input name="image1" type="hidden" id="image1" value="<%=rs2("pic1")%>"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">广告1<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image1" frameborder="0" scrolling="No" width="300" height="25"></iframe> 
      (171*109)</td>
    </tr>			
	    <input name="image2" type="hidden" id="image2" value="<%=rs2("pic2")%>"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">广告2<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image2" frameborder="0" scrolling="No" width="300" height="25"></iframe> 
      (171*109)</td>
    </tr>	
	    <input name="image3" type="hidden" id="image3" value="<%=rs2("pic3")%>"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">广告3<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image3" frameborder="0" scrolling="No" width="300" height="25"></iframe> 
      (171*109)</td>
    </tr>	
		    <input name="image4" type="hidden" id="image4" value="<%=rs2("pic4")%>"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">广告4<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image4" frameborder="0" scrolling="No" width="300" height="25"></iframe> 
      (171*109)</td>
    </tr>		
		    <input name="image5" type="hidden" id="image5" value="<%=rs2("pic5")%>"/>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">广告5<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <iframe id="1" src="upfile1.asp?path=product&name=image5" frameborder="0" scrolling="No" width="300" height="25"></iframe> 
      (171*109)</td>
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
		sql="select * from guanggao where id=1 "
		rs1.open sql,conn,1,3
		rs1("pic1")=request("image1")'标题
		rs1("pic2")=request("image2")'标题
		rs1("pic3")=request("image3")'标题
		rs1("pic4")=request("image4")'标题
		rs1("pic5")=request("image5")'标题
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
	Response.Write "<script>alert('您已经成功修改');location='guangao.asp'</script>"
end if	
%>
</body>
</html>


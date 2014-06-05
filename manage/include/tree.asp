<!-- saved from url=(0040)http://127.0.0.1/manage/include/tree.htm -->
<HTML><HEAD><TITLE></TITLE>
<%
set conn = server.CreateObject("adodb.connection")
conn.ConnectionTimeout = 25
conn.Open "Provider=Microsoft.JET.OLEDB.4.0;" &"Data Source=" & Server.MapPath("../../data/db.mdb") & ";Jet OLEDB:Database Password="
%>
<!--#include file="../Test.asp"-->
<STYLE type=text/css>A:link {
	FONT-SIZE: 9pt; COLOR: #000000; TEXT-DECORATION: none
}
A:active {
	FONT-SIZE: 9pt; COLOR: #000000; TEXT-DECORATION: none
}
A:visited {
	FONT-SIZE: 9pt; COLOR: #000000; TEXT-DECORATION: none
}
.menu {
	FONT-SIZE: 12px; CURSOR: hand
}
A:hover {
	FONT-SIZE: 9pt; COLOR: #000000
}
BODY {
	SCROLLBAR-FACE-COLOR: #e0e9f8; SCROLLBAR-HIGHLIGHT-COLOR: #aaaaaa; SCROLLBAR-SHADOW-COLOR: #aaaaaa; SCROLLBAR-3DLIGHT-COLOR: #e0e9f8; SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #ffffff; SCROLLBAR-DARKSHADOW-COLOR: #e0e9f8; body: #
}
#startingMsg {
	FONT-SIZE: 12px; LEFT: auto; COLOR: #000000; FONT-FAMILY: "Arial"; POSITION: absolute; TOP: auto; HEIGHT: auto
}
</STYLE>

<SCRIPT language=JavaScript>
<!--
if (document.images){
t_sub=new Image
tsub=new Image
t_plus=new Image
tplus=new Image
t_open=new Image
t_close=new Image

t_sub.src="../images/t_sub.gif";
tsub.src="../images/sub.gif";
t_plus.src="../images/t_plus.gif"
tplus.src="../images/plus.gif"
t_open.src="../images/open.gif"
t_close.src="../images/close.gif"
}
//
function showdiv(num){
  var imagename
  var name
  var total
  if (document.all("subtree"+num).style.display=="none" ){
                          document.all.item("subtree"+num).style.display=""
						  		if(num>10000000){
									document.all.item("image"+num).src=tsub.src
									}else{document.all.item("image"+num).src=t_sub.src}
						  document.all.item("img"+num).src=t_open.src
						  }else{
						 document.all.item("subtree"+num).style.display="none"
						 	if(num>10000000){
							document.all.item("image"+num).src=tplus.src
							}else{document.all.item("image"+num).src=t_plus.src}
						 
						 document.all.item("img"+num).src=t_close.src
						  }
  }
  //tips


var browser;
if (document.all) {
		layerRef='document.all.'
		styleRef='.style.'
		changeMessages=".innerHTML=messages[num]"
		closeit=""
		browser=true
}
else
{
   alert("此效果在Netscape浏览器中不能实现！");
}
function mover(num){

	if(browser){
		eval(layerRef+'startingMsg'+changeMessages)
		eval(layerRef+'startingMsg'+closeit);
	}

}

function mout(num){
	if(browser){
		eval(layerRef+'startingMsg'+changeMessages);
		eval(layerRef+'startingMsg'+closeit);
	}
}
-->
</SCRIPT>

<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></HEAD>
<BODY bgColor=#e0e9f8 leftMargin=7 topMargin=0>
<%
if request.Cookies("username")="" then
Response.Write("<script language=javascript>window.open('../system/login.asp','mainFrame','');</script>")
end if
%>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
  <TBODY>
  <TR>
    <TD><BR><SPAN id=sdf style="FONT-SIZE: 12px; COLOR: #ff6600"><IMG 
      src="../images/close.gif" 
      align=absMiddle>&nbsp;<STRONG>网站后台管理系统</STRONG><BR>
    <!----------------------信息发布(3)--------------------------------->
	  </SPAN>

	
	  		<SPAN 
      class=menu id=zhugan10><IMG src="../images/t_subl.gif" 
      align=absMiddle><IMG src="../images/close.gif" align=absMiddle border=0> 
      <SPAN onmouseover=mover(17)><A 
      href="../jiameng.asp" 
      target=mainFrame>招聘信息</A></SPAN></SPAN><BR>
	  
	  <SPAN   class=menu id=zhugan3 onclick=showdiv(3)><IMG src="../images/t_plus.gif" 
      align=absMiddle name=image3><IMG src="../images/close.gif" 
      align=absMiddle border=0 name=img3> <FONT color=#000000><SPAN 
      onmouseover=mover(2)>新闻中心</SPAN></FONT></SPAN><BR>
      <DIV id=subtree3 style="DISPLAY: none">
	  <IMG src="../images/line_05.gif" 
      align=absMiddle><A href="../news_class.asp" 
      target=mainFrame>类别管理</A><BR>
	  <IMG src="../images/line_05.gif" 
      align=absMiddle><A href="../news_add.asp" 
      target=mainFrame>加入信息</A><BR>
	  <IMG src="../images/line_05.gif" 
      align=absMiddle><A href="../news_list.asp" 
      target=mainFrame>管理信息</A>
</DIV>
     	  <SPAN 
      class=menu id=zhugan6131 onclick=showdiv(6131)><IMG src="../images/t_plus.gif" 
      align=absMiddle name=image6131><IMG src="../images/close.gif" 
      align=absMiddle border=0 name=img6131> <SPAN 
      onmouseover=mover(6131)>友情链接</SPAN></SPAN><BR>
      <DIV id=subtree6131 style="DISPLAY: none">
      <IMG src="../images/line_05.gif" 
      align=absMiddle><A href="../link.asp" 
      target=mainFrame>管理信息</A>
	  </DIV>
  <SPAN 
      class=menu id=zhugan6 onclick=showdiv(6)><IMG src="../images/t_plus.gif" 
      align=absMiddle name=image6><IMG src="../images/close.gif" 
      align=absMiddle border=0 name=img6> <SPAN 
      onmouseover=mover(5)>图片资讯管理</SPAN></SPAN><BR>
      <DIV id=subtree6 style="DISPLAY: none">
      <IMG src="../images/line_05.gif" 
      align=absMiddle><A href="../product_class.asp" 
      target=mainFrame>类别管理</A><BR>
      <IMG src="../images/line_05.gif" 
      align=absMiddle><A href="../product_add.asp" 
      target=mainFrame>添加资讯</A><BR>
	  <IMG src="../images/line_05.gif" 
      align=absMiddle><A href="../product_list.asp" 
      target=mainFrame>管理资讯</A><BR>
	  </DIV>

	  		
	  <!---------------------系统维护---------------------------><SPAN 
      class=menu id=zhugan18 onclick=showdiv(18)><IMG 
      src="../images/t_plus.gif" align=absMiddle name=image18><IMG 
      src="../images/close.gif" align=absMiddle border=0 name=img18> <SPAN 
      onmouseover=mover(18)>管理员管理</SPAN></SPAN><BR>
      <DIV id=subtree18 style="DISPLAY: none"><IMG src="../images/line_05.gif" 
      align=absMiddle> <A 
      href="../admin.asp?action=add"  target=mainFrame>添加</A><BR>
	  <IMG src="../images/line_05.gif" 
      align=absMiddle> <A 
      href="../admin.asp?action=edit"  target=mainFrame>修改密码</A><BR>
	    <IMG src="../images/line_05.gif" 
      align=absMiddle> <A 
      href="../admin.asp?action=list"  target=mainFrame>管理</A><BR></DIV>
	  <!----------------- 退出----------------------><SPAN 
      class=menu id=zhugan10><IMG src="../images/t_subl.gif" 
      align=absMiddle><IMG src="../images/close.gif" align=absMiddle border=0> 
      <SPAN onmouseover=mover(17)><A 
      href="../exit.asp" 
      target=_top>退出</A></SPAN></SPAN><BR></TD></TR>
  <TR>
    <TD class=menu>&nbsp;</TD></TR>
  <TR>
    <TD class=menu><STRONG><FONT color=#ff0000><IMG height=32 
      src="../images/tips_title.gif" width=156></FONT></STRONG></TD></TR>
  <TR>
    <TD class=menu>
      <DIV 
      id=startingMsg>通过网站后台管理系统,能够使客户完全在线管理网站.</DIV></TD></TR></TBODY></TABLE></BODY></HTML>

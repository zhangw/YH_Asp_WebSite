<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta content="IE=EmulateIE7" http-equiv="X-UA-Compatible" /> 
<title>媒体新闻</title>
</head>
<script src="../js/jquery-1.4.1.js" language="javascript"></script>
<script language="javascript" src="../js/show_page.js"></script>
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/str.asp"-->
<!--#include file="../inc/fenye.asp"-->
<body>
	<div class="all">
	<div class="content_ny">
	    <table width="984px" border="0" cellpadding="0" cellspacing="0">
		   <tr>
		      <td valign="top" align="right" style="width:220px; height:auto; overflow:hidden;">
<%
smallclassname=trim(request("smallclassname"))
if smallclassname="" then smallclassname="集团新闻"
%>
			  </td>
			  <td style="width:765px; height:auto; overflow:hidden;" align="left" valign="top">
			    <div style="width:765px; height:auto; overflow:hidden;">
				  <table width="765px" border="0" cellpadding="0" cellspacing="0">
				     <tr>
					    <td height="13px" style="background-image:url(images/about_06.jpg); background-repeat:no-repeat; background-position:left top;"></td>
					 </tr>
					 <tr>
					    <td height="440px" valign="top" style="background-image:url(images/about_37.jpg); background-repeat:repeat-y; background-position:left;">
					      <div style="width:705px; text-align:center; height:auto; padding:15px 30px;">
						       <table style="width:700px" border="0" cellpadding="0" cellspacing="0">
								    <tr>
									   <td width="15" height="30px" align="right" style="border-bottom:1px dashed #c3c3c3;" valign="middle"><img src="images/about_09.jpg" width="15" height="19" border="0"></td>
									   <td width="333" align="left" style="border-bottom:1px dashed #c3c3c3;" valign="middle">
									             <font style="font-size:15px; font-family:'微软雅黑'; color:#666; font-weight:bold;">
						              <%if request("smallclassname")<>"" then%><%=request("smallclassname")%><%else%><%response.Write("集团新闻")%><%end if%></font>										</td>
									   <td width="352" style="border-bottom:1px dashed #c3c3c3;" align="right" valign="middle">
									   <font style="font-size:12px; font-family:'微软雅黑'; color:#acacac;">
							            <a href="index.asp" style="font-size:12px; font-family:'微软雅黑'; color:#acacac; text-decoration:none;">首页</a> | <a href="news.asp" style="font-size:12px; font-family:'微软雅黑'; color:#acacac; text-decoration:none;">新闻中心</a> |  
						               <%if request("smallclassname")<>"" then%><%=request("smallclassname")%><%else%><%response.Write("集团新闻")%><%end if%></font></td>
								    </tr>
									<tr>
									   <td  colspan="3" align="center" valign="top">
									       <div style="width:680px; text-align:right; padding:20px 0px 20px 15px; height:auto;">
										      <div style="width:680px; height:40px; overflow:hidden;">
											    <form action="search.asp?smallclassname=<%=request("smallclass")%>" name="myform" style="padding:0px; margin:0px;" method="post"> <table  width="680px" height="40" border="0" cellpadding="0" cellspacing="0">
												     <tr>
													     <td width="8px" style="background-image:url(images/news1_06.jpg); background-repeat:no-repeat; background-position:right bottom;"></td>
													     <td width="660px" style="background-image:url(images/news1_08.jpg); background-repeat:repeat-x; background-position:bottom;"> 
														    <table width="660px" height="40px" border="0" cellpadding="0px" cellspacing="0px">
															   <tr>
															      <td width="245">&nbsp;</td>
															      <td width="36" align="center" valign="middle"><img src="images/news1_13.jpg" width="28" height="23"></td>
															      <td width="37" align="center" valign="middle"><font style="font-size:12px; font-weight:400; font-family:'微软雅黑';">搜索</font></td>
															      <td width="84" valign="middle" align="center">
																  <select style="width:80px;"  size="1"  name="smallclass">
																     <option selected="selected" value="集团新闻">集团新闻</option>
																	 <option value="行业新闻">行业观点</option>
																  </select></td>
															      <td width="170"><input type="text" name="title" style=" height:16px;border:1px solid #666666;" /></td>
															      <td width="63" valign="middle" style="padding-left:8px;" align="left"><input type="image" src="images/news1_15.jpg" width="55" height="18"></td>
															      <td width="25">&nbsp;</td>
															   </tr>
															</table>
														
														 </td>
													     <td width="12px" style="background-image:url(images/news1_10.jpg); background-repeat:no-repeat; background-position:left bottom;"></td>
												     </tr>
												 </table>
												  </form>
											  </div>     
										     <div style="width:680px; height:45px; overflow:hidden; padding-top:20px;">
											     <table  width="680px" height="45px" border="0" cellpadding="0" cellspacing="0">
												     <tr>
													     <td width="10" style="background-image:url(images/news1_14.jpg); background-repeat:no-repeat; background-position:right bottom;"></td>
													     <td align="left" valign="middle" width="547" style="background-image:url(images/news1_31.jpg); background-repeat:repeat-x; background-position:bottom; "><div style="font-size:14px; font-weight:500; color:#ed9898;  font-family:微软雅黑';">新闻标题</div></td>
													     <td width="112" align="center" valign="middle" style="background-image:url(images/news1_31.jpg); background-repeat:repeat-x; background-position:bottom;"><img src="images/news1_36.jpg" width="89" height="22" border="0"></td>
													     <td width="11" style=" background-image:url(images/news1_33.jpg); background-repeat:no-repeat; background-position:left bottom;"></td>
												     </tr>
												 </table>
										     </div> 
											 <div style="width:660px; height:auto; padding:10px 10px 0px 10px; overflow:hidden;">
											     <table  width="660px" height="auto" border="0" cellpadding="0" cellspacing="0">
												 <%
												    set rs_classid=server.CreateObject("adodb.recordset")
													sql_classid="select * from newsid where title='"&smallclassname&"' order by newsclass desc"
													rs_classid.open sql_classid,conn,1,1
													if rs_classid.bof and rs_classid.eof then
													  response.Write("<script language=""JavaScript"">alert(""暂无消息"");</script>")
												    else
													  page=clng(request("page")) 
													  set rs_news=server.CreateObject("adodb.recordset")
													  sql_news="select * from news where newsid in("&rs_classid("id")&") order by id desc"
													  rs_news.open sql_news,conn,1,1
													  if rs_news.eof and rs_news.bof  then  
													      response.Write("暂无消息~！")
													  else
													  rs_news.PageSize=8
                                                      if page=0 then page=1 
                                                      pages=rs_news.pagecount
                                                      if page > pages then page=pages
                                                      rs_news.AbsolutePage=page 
													  for i=1 to 8
													  
												 %>
												     <tr>
													     <td width="568" align="left" style="border-bottom:1px dashed #f2f2f2" height="30px">
														     <a href="show_news.asp?id=<%=rs_news("id")%>&smallclassname=<%=smallclassname%>" style="font-size:12px; text-decoration:none; font-weight:300; color:#5c5c5c; font-family:'微软雅黑';"><%=left(rs_news("title"),35)%></a>														 </td>
													   <td width="92"  align="left" style="border-bottom:1px  dashed #f2f2f2">&nbsp;<font style="font-size:12px; font-weight:300; color:#5c5c5c; font-family:'微软雅黑';"><%=replace(formatdatetime(rs_news("time"),2),"/",".")%></font></td>
												     </tr>
													 
													 <%
													      rs_news.movenext
														  if rs_news.eof then exit for 
														  next
														  end if
													 %>
														 <tr page="">
													    <td colspan="2"><span  style="float:left; width:565px; text-align:right; padding-top:15px; font-size:12px; font-family:'微软雅黑';"><Script Language=JavaScript>
ShowoPage("","","页:<font color='red'>","</font>/","<font color='red'>","</font>页&nbsp;","&nbsp;每页<font color='red'>","</font>条&nbsp;","&nbsp;共<font color='red'>","</font>条&nbsp;&nbsp;","<font  color='black'>首页</font>","<font color='black'>上一页</font>","<font color='black'>下一页</font>","<font   color='black'>尾页</font>","&nbsp;跳转:","<font color='red'>","</font>","[<font color='red'>","</font>]","","","&nbsp;","&nbsp;",<%=(rs_news.recordcount)%>,<%=rs_news.pagesize%>,2)
</Script>
</span></td>
													 </tr>	  
												     <%
					                                   rs_news.close
					                                   set rs_news=nothing
													 end if
					                                         rs_classid.close
															 set rs_classid=nothing
												     %>
												
												 </table>
										     </div> 
									       </div>									  
									   </td>
								    </tr>
								</table>
						  </div>
						</td>
					 </tr>
				  </table>
				</div>
			  </td>
		   </tr>
		</table>
	</div><!--content-->
</div>
</body>
</html>

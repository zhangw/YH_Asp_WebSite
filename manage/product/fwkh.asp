<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta content="IE=EmulateIE7" http-equiv="X-UA-Compatible" /> 
<title>乔恩广告传媒有限公司</title>
<link rel="stylesheet" type="text/css" charset="UTF-8"  href="./layout.css"  />
<link rel="stylesheet" type="text/css" charset="UTF-8"  href="./style.css" />
</head>
<script src="js/jquery-1.4.1.js" language="javascript"></script>
<script language="javascript">
<!--
//-->
</script>
<body>
<div class="all">
    <!--#include file="top2.asp"-->
	<div class="banner_ny">
	    <div style="width:984px; padding:0px;">
		    <img src="images/product_03.jpg" width="984" height="135">
		</div><!--banner-->
	</div>
	<div class="content_ny">
	    <table width="984px" border="0" cellpadding="0" cellspacing="0">
		   <tr>
		      <td valign="top" align="right" style="width:220px; height:auto; overflow:hidden;">
			           <!--#include file="left_cp.asp"-->
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
									   <td width="21" height="30px" align="right" style="border-bottom:1px dashed #c3c3c3;" valign="middle"><img src="images/about_09.jpg" width="15" height="19" border="0"></td>
									   <td width="320" align="left" style="border-bottom:1px dashed #c3c3c3;" valign="middle">
									             <font style="font-size:15px; font-family:'微软雅黑'; color:#666; font-weight:bold;"><%if request("smallclassname")<>"" then%><%=request("smallclassname")%><%else%><%response.Write("成功案例")%><%end if%></font>										</td>
									   <td width="343" style="border-bottom:1px dashed #c3c3c3;" align="right" valign="middle">
									   <font style="font-size:12px; font-family:'微软雅黑'; color:#acacac;">
							            <a href="index.asp" style="font-size:12px; font-family:'微软雅黑'; color:#acacac; text-decoration:none;">首页</a> | <a href="product.asp" style="font-size:12px; font-family:'微软雅黑'; color:#acacac; text-decoration:none;">产品与服务</a> |  <%if request("smallclassname")<>"" then%><%=request("smallclassname")%><%else%><%response.Write("成功案例")%><%end if%></font></td>
								    </tr>
									<tr>
									   <td  colspan="3" align="center" valign="top">
									       <div style="width:680px; text-align:right; padding:20px 0px 20px 20px; height:auto;">
										        <table width="680px" border="0" height="360px" cellpadding="0" cellspacing="0">
												   <tr>
												      <td width="680px"  align="left" valign="top" height="360px" >
													  <div style="width:660px; text-align:left; height:420px; overflow:hidden;">
                                                                   <table  height="122" border="0" cellpadding="0" cellspacing="8px">
																  <% 
						     set rs_productid=server.CreateObject("adodb.recordset")
							 sql_productid="select * from linkclass order by id desc"
							 rs_productid.open sql_productid,conn,1,1
							 if rs_productid.bof and rs_productid.eof  then  
							     response.Write("<script language=""JavaScript"">alert(""暂无分类，请从后台添加"");history.go(-1);</script>")
							 else
							     productclassid=rs_productid("id")	
							 end if
							 rs_productid.close
							 set rs_productid=nothing
						                                            page=clng(request("page")) 
                                                                    set rs_product=server.CreateObject("adodb.recordset")
																	sql_product="select * from link  where class in ("&productclassid&") order by id desc"
																	rs_product.open sql_product,conn,1,1
																	rs_product.PageSize=15
																    if page=0 then page=1 
                                                                    pages=rs_product.pagecount
                                                                    if page > pages then page=pages
                                                                    rs_product.AbsolutePage=page 
																	if rs_product.eof and rs_product.bof  then response.Write("暂无信息~！")
																	for i=1 to 3
                                                                   %>
																     <tr height="130px" valign="top">
																	     <%
																		    for j=1 to 5
																		 %>
																	   <td align="center" width="120px" height="102" valign="middle"><a href="<%=rs_product("link")%>" style="text-decoration:none;" target="_blank" title="<%=rs_product("title")%>"><img src="manage/product/<%=rs_product("pic")%>"  alt="<%=rs_product("title")%>" class="bordered" border="0" /></a>																       </td>
																	   <%
																		   rs_product.movenext
																		   if rs_product.eof then exit for
																		   next
																		%>
																	 </tr>
																	 <%
																		   if rs_product.eof then exit for
																		   next
																		 
																     %>
													    </table>
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
					 <tr>
					    <td height="10px" style="background-image:url(images/about_39.jpg); background-repeat:no-repeat; background-position:left bottom;"></td>
					 </tr>
				  </table>
				</div>
			  </td>
		   </tr>
		</table>
	</div><!--content-->
	<!--
	<div class="yqlj_index">
	      <div style="width:980px; margin:0px auto; height:71px; overflow:hidden; padding-top:15px;">
	          <div id="demo" style="overflow:hidden; width:968px; height:81px; color:#ffffff;">
                <table cellpadding="0" cellspacing="0" border="0">
                  <tr>
                    <td id="demo1" valign="top" align="center"><table cellpadding="2" cellspacing="2" border="0">
                        <tr align="center">
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_02.gif" width="95" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_03.gif" width="85" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_04.gif" width="79" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_05.gif" width="86" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_06.gif" width="80" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_07.gif" width="86" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_08.gif" width="82" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_09.gif" width="96" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_10.gif" width="73" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_11.gif" width="73" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_12.gif" width="91" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_07.gif" width="86" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                          <td><a href="http://www.baidu.com" target="_blank"><img src="images/1_04.gif" width="79" height="71"  border="0" style="cursor:pointer;"  /></a></td>
                        </tr>
                    </table></td>
                    <td id="demo2" valign="top"></td>
                  </tr>
                </table>
				</div>
            </div>
	          <script> 
  var speed=15//速度数值越大速度越慢 
  demo2.innerHTML=demo1.innerHTML 
  function Marquee(){ 
  if(demo.scrollLeft<=0) 
  demo.scrollLeft+=demo2.offsetWidth 
  else{ 
  demo.scrollLeft-- 
  } 
  } 
  var MyMar=setInterval(Marquee,speed) 
  demo.onmouseover=function() {clearInterval(MyMar)} 
  demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)} 
  </script> 
	</div>--><!--滚动链接-->
	<!--#include file="foot.asp"-->
</div>
</body>
</html>

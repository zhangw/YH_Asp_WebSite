<!--#include file="conn.asp"-->
<!--#include file="../admin.asp"-->
<%
	if  instr(session("abi"),"b") = 0 or isnull(instr(session("abi"),"b")) then
		response.Write "你没有权限操作此功能..."
		response.End()
	end if
%>
<!-- #include file="../Inc/Head.asp" -->
<!-- #include file="../Inc/fenye.asp" -->

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top">
	 <br> <strong><br></strong>  
	 <table width="700" border="0">
  		<tr>
    		<td width="671">
			<fieldset style="width:240px;height:50px; border:#9CB8FA 1px solid;">
			<legend style="font-weight:800;">管理选项</legend>
			<br />
			<a href="news.asp?action=add" style="color: #FF0000;"> &nbsp;添加信息</a>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="news.asp?action=list" style="color: #FF0000;">查看信息</a><br />
			<br />			
			</fieldset>
		  </td>
  		</tr>
	</table>
	<br />
	 <% if request("action") = "list" then%>
	  <table width="700" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td width="817" height="23" class="back_southidc"> <div align="center"><strong>信 息 列 表</strong></div></td>
        </tr>
        <tr class="tr_southidc"> 
          <td> 
	  <table width="100%" border="0" cellspacing="2" bgcolor="#6A87BD">
  <tr>
    <td width="41" height="25" align="center" bgcolor="#A4B6D7"><span style="font-weight: bold">序号</span></td>
    <td width="173" align="center" bgcolor="#A4B6D7"><span style="font-weight: bold">名称</span></td>
    <td width="148" align="center" bgcolor="#A4B6D7"><span style="font-weight: bold">缩写</span></td>
    <td width="134" height="25" align="center" bgcolor="#A4B6D7"><span style="font-weight: bold">文件名</span></td>
    <td width="176" height="25" align="center" bgcolor="#A4B6D7"><span style="font-weight: bold">操作</span></td>
  </tr>
	<%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select * from hy_news"
		rs.open sql,conn,1,3
	%>
  <%if rs.eof then%>
  <tr>
    <td colspan="5" align="center" bgcolor="#ECF5FF">暂且没有数据</td>
  </tr>
  <%else%>
	<%
  		rs.PageSize=20
		page=int(request("page"))
		if page="" or page<=0 then 
			page=1
		end if
		all=rs.pagecount
		rs.absolutepage=page
	%>
	<% for i=1 to rs.PageSize %>
  <tr class="fclass" style="display:block">
    <td align="center" bgcolor="#ECF5FF"><%=i%></td>
    <td align="center" bgcolor="#ECF5FF"><%=rs("news_title")%></td>
    <td align="center" bgcolor="#ECF5FF"><%=rs("news_content")%></td>
    <td align="center" bgcolor="#ECF5FF"><%=rs("news_author")%></td>
    <td align="center" bgcolor="#ECF5FF"><a href="news.asp?action=edit&id=<%=rs("id")%>">【编辑】</a> <a href="../sgzs/sgzs1.asp?action=list&pid=<%=rs("id")%>">【小分类】</a> 
	<a href="?true=del&id=<%=rs("id")%>" onClick="return isDel();">【删除】</a></td>
  </tr>
	<%
		rs.movenext
		if rs.eof then exit for	
	    next
		url = url&"news.asp?action=list&"
	%>
  <tr>
    <td colspan="5" align="right" bgcolor="#A4B6D7"><%movepage rs.recordcount,page,all,url %>&nbsp;</td>
  </tr>
	<%end if%>
</table>
</td>
        </tr>
      </table>
<%
	rs.close
	set rs = nothing
%>
<%end if%>
<!----------------------一新闻添加------------------------------->
<%if request("action") = "add" then%>
<form name="formm" action="?action=aa" method="post" onSubmit="return check_content()">
<table width="700" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>信 息 添 加<br>
              </strong></div></td>
        </tr>
        <tr class="tr_southidc"> 
          <td><div align="center"> 
<table width="100%" border="0" cellspacing="2" bgcolor="#6A87BD">
  
  
  <tr>
    <td width="79" align="right" bgcolor="#B9DAFF"><span style="font-weight: bold">名称：</span></td>
    <td width="460" bgcolor="#ECF5FF"><input name="title" type="text" id="title" size="50"></td>
    <td width="97" bgcolor="#ECF5FF">&nbsp;<span style="color: #FF0000">*</span></td>
  </tr>
    <tr>
    <td width="79" align="right" bgcolor="#B9DAFF"><span style="font-weight: bold">缩写：</span></td>
    <td width="460" bgcolor="#ECF5FF"><input name="content" type="text" id="content" size="50"></td>
    <td width="97" bgcolor="#ECF5FF">&nbsp;<span style="color: #FF0000">*</span></td>
  </tr>
    <tr>
    <td width="79" align="right" bgcolor="#B9DAFF"><span style="font-weight: bold">文件名：</span></td>
    <td width="460" bgcolor="#ECF5FF"><input name="mc" type="text" id="mc" size="50"></td>
    <td width="97" bgcolor="#ECF5FF">&nbsp;<span style="color: #FF0000">*</span></td>
  </tr>
  <tr>
    <td colspan="3" align="center" bgcolor="#A4B6D7">
		<input type="submit" name="Submit" value=" 提 交 ">&nbsp;&nbsp;
		<input type="reset" name="Submit2" value=" 取 消 ">&nbsp;&nbsp;
		<input type="button" name="Submit3" value=" 返 回 " onClick="javascript:history.back();">	</td>
    </tr>
</table>
</td>
</tr>
</table>
</form>
<%end if%>
<!----------------------新闻编辑------------------------------->
<%if request("action") = "edit" then%>
<%
	set rs = server.CreateObject("adodb.recordset")
		sql = "select * from hy_news where id="&request("id")
		rs.open sql,conn,1,3
%>
<form name="formm" action="?action=editsave&id=<%=request("id")%>" method="post" onSubmit="return check_content()">
<table width="700" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>信 息 编 辑<br>
              </strong></div></td>
        </tr>
        <tr class="tr_southidc"> 
          <td><div align="center"> 
<table width="100%" border="0" cellspacing="2" bgcolor="#6A87BD">
  
  
  <tr>
    <td width="79" align="right" bgcolor="#B9DAFF"><span style="font-weight: bold">名称：</span></td>
    <td width="460" bgcolor="#ECF5FF"><input name="title" type="text" id="title" value="<%=rs("news_title")%>" size="50"></td>
    <td width="97" bgcolor="#ECF5FF">&nbsp;<span style="color: #FF0000">*</span></td>
  </tr>
    <tr>
    <td width="79" align="right" bgcolor="#B9DAFF"><span style="font-weight: bold">缩写：</span></td>
    <td width="460" bgcolor="#ECF5FF"><input name="content" type="text" id="content" value="<%=rs("news_content")%>" size="50"></td>
    <td width="97" bgcolor="#ECF5FF">&nbsp;<span style="color: #FF0000">*</span></td>
  </tr>
    <tr>
    <td width="79" align="right" bgcolor="#B9DAFF"><span style="font-weight: bold">文件名：</span></td>
    <td width="460" bgcolor="#ECF5FF"><input name="mc" type="text" id="mc" value="<%=rs("news_author")%>" size="50"></td>
    <td width="97" bgcolor="#ECF5FF">&nbsp;<span style="color: #FF0000">*</span></td>
  </tr>
  <tr>
    <td colspan="3" align="center" bgcolor="#A4B6D7">
		<input type="submit" name="Submit" value=" 提 交 ">&nbsp;&nbsp;
		<input type="reset" name="Submit2" value=" 取 消 ">&nbsp;&nbsp;
		<input type="button" name="Submit3" value=" 返 回 " onClick="javascript:history.back();">		</td>
    </tr>
</table>
</td>
</tr>
</table>
</form>
<%
	rs.close
	set rs = nothing
%>
<%end if%>
<span id="hint"></span>
    </td>
  </tr>
</table>
<!-- #include file="../Inc/Foot.asp" -->
<!-------------------------一级新闻添加----------------------------->
<%if request("action") = "aa" then%>
	<%
		d3 = trim(request("title"))
		
		d7 = trim(request("content"))
		
		if d3 = "" then
			response.Write "<script>alert('请输入名称...');history.back();</script>"
			response.End()
		end if
		if d7 = "" then
			response.Write "<script>alert('请输入缩写...');history.back();</script>"
			response.End()
		end if
		if  trim(request("mc"))="" then
			response.Write "<script>alert('请输入文件名...');history.back();</script>"
			response.End()
		end if
		set rs = server.CreateObject("adodb.recordset")
		sql = "select * from hy_news"
		rs.open sql,conn,1,3
		rs.addnew
		
		rs("news_title") = d3
		rs("news_author") = trim(request("mc"))
		rs("news_content") = d7
		
		rs.update
		rs.close
		set rs = nothing
		response.Write "<script>alert('添加成功...');location.reload('news.asp?action=list');</script>"
	%>
<%end if%>
<!-------------------------一级新闻修改----------------------------->
<%if request("action") = "editsave" then%>
	<%
		
		d3 = trim(request("title"))
		
		d7 = trim(request("content"))
		if d3 = "" then
			response.Write "<script>alert('请输入名称...');history.back();</script>"
			response.End()
		end if
		if d7 = "" then
			response.Write "<script>alert('请输入缩写...');history.back();</script>"
			response.End()
		end if
		if trim(request("mc")) = "" then
			response.Write "<script>alert('请输入文件名...');history.back();</script>"
			response.End()
		end if
		set rs = server.CreateObject("adodb.recordset")
		sql = "select * from hy_news where id="&request("id")
		rs.open sql,conn,1,3
		
		rs("news_title") = d3
		rs("news_author") = trim(request("mc"))
		rs("news_content") = d7
		
		rs.update
		rs.close
		set rs = nothing
		
		
		response.Write "<script>alert('修改成功...');location.reload('news.asp?action=list');</script>"
	%>
<%end if%>
<!-------------------------新闻删除----------------------------->
<%if request("true") = "del" then%>
	<%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select * from hy_news where id="&request("id")
		rs.open sql,conn,1,3
		rs.delete
		rs.update
		rs.requery
		rs.close
		set rs = nothing
		
		
		response.Write "<script>alert('删除成功...');location.reload('news.asp?action=list');</script>"
	%>
<%end if%>

<html>
<head>
<%if request.cookies("username")=""  Then
  response.Redirect "login.asp"
end if
%>

<title>智能网站II - 后台管理-2.2</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head>

<frameset rows="79,*" cols="*" frameborder="NO" border="0" framespacing="0">
  <frame src="include/header.htm" name="topFrame" scrolling="NO" noresize >
  <frameset id=f2 name=f2 cols="190,*" frameborder="NO" border="0" framespacing="0">
    <frame src="include/tree.asp" name="leftFrame" scrolling="auto" noresize>
        <frameset rows="*" cols="8,*" framespacing="0" frameborder="NO" border="0">
      		<frame src="include/button.htm" name="leftFrame" scrolling="NO" noresize>
			<frame src="include/start.htm" name="mainFrame">
  		</frameset>
</frameset>
 </frameset>
<noframes><body>

</body></noframes>
</html>
</html>
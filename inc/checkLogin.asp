<!-- #include file="conn.asp"-->
<!-- #include file="inc/function.asp"-->
<!-- #include file="inc/md5.asp"-->
<%
	user = trim(request("UserName"))
	pwd = trim(request("Password"))
	code = trim(request("Code"))
	if user = "" then
		call errview(1,"login.asp")
		response.End()
	end if
	if pwd = "" then
		call errview(2,"login.asp")
		response.End()
	end if
	if code = "" then
		call errview(4,"login.asp")
		response.End()
	end if
	if cint(code) <> cint(Session("GetCode")) then
		call errview(5,"login.asp")
		response.End()
	end if
	'¼ì²â·Ç·¨×Ö·û
	if reg(user,1) then
		call errview(6,"login.asp")
		response.End()
	end if
	if reg(pwd,2) then
		call errview(7,"login.asp")
		response.End()
	end if
	set rs = server.CreateObject("adodb.recordset")
		sql = "select * from admin"
		rs.open sql,conn,1,1
		do while not rs.eof
			if rs("admin_user") = user and rs("admin_pwd") = md5(pwd) then
				session("login") = "<ok>"
				session("AdminName") = user
				session("AdminPassword") = pwd
				session("abi") = rs("admin_abi")
				rs.close
				set rs = nothing
				response.Redirect "default.asp"
				exit do
			end if
			rs.movenext
		loop
		rs.close
		set rs = nothing
		call errview(8,"login.asp")
	
'LoadPicture(picturename)
%>
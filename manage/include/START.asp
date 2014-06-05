<%if request.cookies("username")=""  Then
  response.Redirect "../system/login.asp"
  else
  response.Redirect "START.htm"
end if
%>

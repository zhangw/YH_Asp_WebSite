<%if request.cookies("username")=""  Then
  response.Redirect "system/login.asp"
end if
%>

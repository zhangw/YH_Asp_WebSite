<!--#include file="connstr.asp"-->
<%
set conn = server.CreateObject("adodb.connection")
conn.ConnectionTimeout = 25
conn.Open "Provider=Microsoft.JET.OLEDB.4.0;" &"Data Source=" & Server.MapPath("../data/db.mdb") & ";Jet OLEDB:Database Password="
 
'请注意及时调用该函数关闭数据库的连接

sub closedb()              '声明函数
	on error resume next'有错误继续输出下面页面
	conn.close             '关闭数据库连接
	set conn=nothing       '清空数据库连接
	err.clear              '清空错误
end sub

Function SafeRequest(ParaName,ParaType)
'--- 传入参数 
'ParaName:参数名称-字符型
'ParaType:参数类型-数字型(1表示以上参数是数字，0表示以上参数为字符)
	Dim ParaValue
	ParaValue=Request(ParaName)
	If ParaType=1 then
		If not isNumeric(ParaValue) then
			response.write "<script language=javascript>alert('输入格式不正确!');"
			response.write "location.href='login.asp';</script>"
			response.end
		End if
	Else
		ParaValue=replace(ParaValue,"'","''")
	End if
	SafeRequest=ParaValue
End function
%>
<%

'防止ＳＱＬ注入式攻击代码。不要删除。谢谢

'对地址栏输入的字符串进行检索，找出非法字符，转向指定页面。
Dim Fy_Url,Fy_a,Fy_x,Fy_Cs(),Fy_Cl,Fy_Ts,Fy_Zx
'---定义部份  头------
Fy_Cl = 3        '处理方式：1=提示信息,2=转向页面,3=先提示再转向
Fy_Zx = "/"    '出错时转向的页面
'---定义部份  尾------
On Error Resume Next
Fy_Url=Request.ServerVariables("QUERY_STRING")
Fy_a=split(Fy_Url,"&")
redim Fy_Cs(ubound(Fy_a))
On Error Resume Next
for Fy_x=0 to ubound(Fy_a)
Fy_Cs(Fy_x) = left(Fy_a(Fy_x),instr(Fy_a(Fy_x),"=")-1)
Next

For Fy_x=0 to ubound(Fy_Cs)
If Fy_Cs(Fy_x)<>"" Then
If Instr(LCase(Request(Fy_Cs(Fy_x))),"'")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"and")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"select")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"update")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"chr")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"delete%20from")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),";")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"insert")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"mid")<>0 Or Instr(LCase(Request(Fy_Cs(Fy_x))),"master.")<>0 Then
Select Case Fy_Cl
  Case "1"
Response.Write "<Script Language=javascript>alert('  出现错误！参数 "&Fy_Cs(Fy_x)&" 的值中包含非法字符串！\n\n  n');window.close();</Script>"
  Case "2"
Response.Write "<Script Language=javascript>location.href='"&Fy_Zx&"'</Script>"
  Case "3"
Response.Write "<Script Language=javascript>alert('  出现错误！参数 "&Fy_Cs(Fy_x)&"的值中包含非法字符串！\n\n');location.href='"&Fy_Zx&"';</Script>"
End Select
Response.End
End If
End If
Next


'判断提交地点。如不是由服务程序提交则返回首页
referer=Cstr(Request.ServerVariables("HTTP_REFERER"))
ser_name=Cstr(Request.ServerVariables("SERVER_NAME"))
'判断浏览器位置
'response.write "\"& referer &"\"& ser_name
'response.end
'if mid(referer,8,len(ser_name))<>ser_name then
 '  response.redirect "/"
'end if
'防止ＳＱＬ注入式攻击代码。不要删除。谢谢
%>

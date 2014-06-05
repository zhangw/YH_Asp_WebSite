<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
								<td class="SheetSelected" vAlign="bottom" width="60">鍚庡彴绠＄悊</td>
								<td valign="bottom"><img src="images/sheet_right.gif" width="25" height="18"></td>
							</tr>
						</table>
					</td>
					<td vAlign="bottom" align="right">
						<table cellSpacing="2" cellPadding="0" width="99%" border="0">
							<tr>
								<td align="right">鎮ㄧ殑浣嶇疆锛?A href="index.asp" target="_top">鍚庡彴绠＄悊</A> &gt;&gt; 
									鐣欒█绠＄悊</td>
						  </tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>

          <%if request("action")="" then%><TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle">浜у搧 淇℃伅鐨勫鍔犮€佷慨鏀瑰拰鍒犻櫎銆?/TD>
                  <TD align="right">聽</TD>
                </TR>
                <TR>
                  <TD>聽</TD>
                  <TD align="right">聽</TD>
                </TR>
              </TBODY>
            </TABLE>
	<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#A4B6D7" class="unnamed1">
    <td width="85" height="20" bgcolor="#e0e9f8" class="STYLE5"><div align="center"><strong>鐢ㄦ埛鍚?/strong></div></td>
    <td width="120" bgcolor="#e0e9f8" class="STYLE5"><div align="center"><strong>濮撳悕</strong></div></td>
    <td width="161" bgcolor="#e0e9f8" class="STYLE5"><div align="center"><strong>鐢佃瘽</strong></div>     </td>
    <td width="196" align="center" bgcolor="#e0e9f8" class="STYLE5"><strong>娉ㄥ唽鏃堕棿</strong></td>
    <td width="93" height="25" bgcolor="#e0e9f8" class="STYLE5"><div align="center"><strong>鎿嶄綔</strong></div></td>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from member " 
	sql=sql & " order by id desc"
	rs.open sql,conn,1,3
if  rs.eof then
 response.Write "娌℃湁璁板綍!"
 response.End()
end if
		rs.pagesize=13
	if Request.QueryString("page") then page=cint(Trim(Request.QueryString("page")))
	if page<1 then
		page=1
	elseif page>rs.pagecount then
		page=rs.pagecount
	end if
	rs.AbsolutePage=page
rowCount = rs.PageSize

	do while not rs.eof and rowcount>0
%>
  <tr bgcolor="#ECF5FF" class="unnamed1">
    <td width="85" height="25" bgcolor="#DDF4FF" class="STYLE5"><%=rs("name")%></td>
	<td width="120" bgcolor="#DDF4FF" class="STYLE5"><div align="center"><%=rs("submitdate")%></div></td>
    <td  bgcolor="#DDF4FF" class="STYLE5"><div align="center"><%=rs("tel")%></div></td>
    <td  bgcolor="#DDF4FF" class="STYLE5"><div align="center"><%=formatdatetime(rs("submitdate"),2)%></div></td>
    <td width="93" bgcolor="#DDF4FF"><div align="center"><a href="?action=edit&id=<%=rs("id")%><%if request("lb")<>"" then%>&lb=<%=request("lb")%><%end if%>" class="hh">鏌ョ湅</a>|<a href="?action=del&id=<%=rs("id")%>" class="hh">鍒犻櫎</a></div></td>
  </tr>
  <%RowCount=RowCount-1
rs.movenext
loop
%>
</table>
<table width="90%"  border="0" align="center" cellspacing="1">
  <tr>
    <td height="20"><div align="center" class="unnamed1">绗?%= page %>椤?nbsp; <a href="?page=1&action=list<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh">棣栭〉</a> &nbsp;鍏?%=rs.PageCount%>椤?nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %>&action=list<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >涓婁竴椤?/a>
            <% else %>
        涓婁竴椤?
        <% end if %>
&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%>&action=list<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >涓嬩竴椤?/a>
        <% else %>
        涓嬩竴椤?
        <% end if %>
&nbsp;
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%>&action=list<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >鏈〉</a>
        <% else %>
        鏈〉
        <% end if %>
&nbsp;鎬绘暟<%= rs.recordcount %>鏉?/div></td>
  </tr>
</table>		
<TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD>聽</TD>
                </TR>
                <TR>
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle">聽 </TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
sql="select * from member where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" />淇敼浼氬憳淇℃伅</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <td width="21%" height="25" align="right">浼氬憳鍚嶏細</td>
                    <td width="79%"><%=rs("admin")%></td>
                  </tr>
                  <tr>
                    <td height="25" align="right">瀵嗙爜锛?/td>
                    <td><input name="password" type="password" id="password" value="<%=rs("password")%>"></td>
                  </tr>
                  <tr>
                    <td height="25" align="right">纭瀵嗙爜锛?/td>
                    <td><input name="password1" type="password" id="password1" value="<%=rs("password")%>"></td>
                  </tr>
                  <tr>
                    <td height="25" align="right">濮撳悕锛?/td>
                    <td><input name="name" type="text" id="name" value="<%=rs("name")%>"></td>
                  </tr>
                  <tr>
                    <td height="25" align="right">鐢佃瘽锛?/td>
                    <td><input name="tel" type="text" id="tel" value="<%=rs("tel")%>"></td>
                  </tr>				  
                  <tr>
                    <td height="25" align="right">閭锛?/td>
                    <td><input name="mail" type="text" id="mail" value="<%=rs("mail")%>"></td>
                  </tr>
                  <tr>
                    <td height="25" align="right">鍦板潃锛?/td>
                    <td><input name="adress" type="text" id="adress" value="<%=rs("adress")%>"></td>
                  </tr>
                  <tr>
                    <td height="25" align="right">绉垎锛?/td>
                    <td><input name="jf" type="text" id="jf" value="<%=rs("jf")%>" /></td>
                  </tr>
                  <tr>
                    <td height="25">&nbsp;</td>
                    <td><input type="submit" name="Submit" value="淇敼"></td>
                  </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
	set rs=server.CreateObject("ADODB.Recordset")
	sql="select * from member where id="&request("id")
	rs.open sql,conn,1,3
	if request("password")="" then
Response.Write "<script>alert('璇疯緭鍏ュ瘑鐮?');history.go(-1)</script>"
response.End()
	end if
	if request("password1")="" then
Response.Write "<script>alert('璇疯緭鍏ョ‘璁ゅ瘑鐮?');history.go(-1)</script>"
response.End()
	end if
		if request("name")="" then
Response.Write "<script>alert('璇疯緭鍏ュ鍚?');history.go(-1)</script>"
response.End()
	end if
		rs("password")=request("password")
		rs("name")=request("name")
		rs("mail")=request("mail")
		rs("tel")=request("tel")
		rs("adress")=request("adress")
		rs("jf")=request("jf")
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		Response.Write "<script>alert('鎮ㄥ凡缁忔垚鍔熶慨鏀?);location='member.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from member where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('鍒犻櫎鎴愬姛');location='member.asp'</script>"
end if
%>		
</body>
</html>

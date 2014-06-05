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
									鍏充簬鎴戜滑</td>
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
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle"> 鍏充簬鎴戜滑鐨勫鍔犮€佷慨鏀瑰拰鍒犻櫎銆?/TD>
                  <TD align="right">聽</TD>
                </TR>
                <TR>
                  <TD>聽</TD>
                  <TD align="right">聽</TD>
                </TR>
              </TBODY>
            </TABLE>
			
            <TABLE width="90%" border="0" align="center" cellPadding="1" cellSpacing="1" bgColor="#c0c0c0">
              <TBODY>
                <TR bgColor="#e0e9f8">
                  <TD align="middle" width="24" height="24">聽 </TD>
                  <TD align="middle"><strong>鏍囬</strong></TD>
                  <TD align="middle" width="50"><STRONG>鎿?浣?/STRONG></TD>
                </TR>
				<%
sql="select * from contact order by id asc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
i=0 
do while not rs.eof
i=i+1
%>
                <TR bgColor="#fffff0">
                  <TD align="middle">&nbsp;</TD>
                  <TD align="left" bgColor="#fffff0"><a href="#"><%=rs("title")%></a></TD>
                  <TD align="middle"><A title="淇敼" href="?action=edit&id=<%=rs("id")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('纭疄瑕佸垹闄ゅ悧'))   href='?action=del&id=<%=rs("id")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="鍒犻櫎"></A></TD>
                </TR>
<%
rs.movenext
loop
%>				
              </TBODY>
            </TABLE>
<TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD>聽</TD>
                </TR>
                <TR>
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="?action=add">鏂板淇℃伅</A>聽 </TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="add" then
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 鏂板淇℃伅</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formc" method="post" action="?action=aa">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="73" height="20" bgcolor="#FFFFFF"><div align="right"><span class="front3">鏍囬</span>锛?/div></td>
      <td align="left" bgcolor="#FFFFFF"><input name="title" type="text" id="title">
      <span class="STYLE5"> (鏍囬璇蜂笉瑕佸ぇ浜?0瀛?</span></td>
    </tr>
	    <tr bgcolor="#A4B6D7">
      <td height="20" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right"><br />
      鍐呭锛?/div></td>
      <td height="250" align="left" bgcolor="#FFFFFF" class="STYLE5"><textarea name="theme" style="visibility:hidden"></textarea><iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe>	</td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
      <td align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="纭畾"></td>
    </tr>
  </table>
</form>
<%end if
if request("action")="aa" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from contact "
	rs.open sql,conn,1,3
	rs.addnew()
	rs("title")=request("title")
	rs("class")=request("class")
	rs("theme")=request("theme")
	rs.update
	rs.close
	set rs=nothing		
	Response.Write "<script>alert('鎮ㄥ凡缁忔垚鍔熸坊鍔?);location='contact.asp'</script>"
end if	
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from contact where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 淇敼淇℃伅</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formc" method="post" action="?action=editsave&id=<%=request("id")%>">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="73" height="20" bgcolor="#FFFFFF"><div align="right"><span class="front3">鏍囬</span>锛?/div></td>
      <td align="left" bgcolor="#FFFFFF"><input name="title" type="text" id="title" value="<%=rs("title")%>">
      <span class="STYLE5"> (鏍囬璇蜂笉瑕佸ぇ浜?0瀛?</span></td>
    </tr>
	    <tr bgcolor="#A4B6D7">
      <td height="20" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right"><br />
      鍐呭锛?/div></td>
      <td height="250" align="left" bgcolor="#FFFFFF" class="STYLE5"><textarea name="theme" style="visibility:hidden"><%=rs("theme")%></textarea><iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe>	</td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
      <td align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="纭畾"></td>
    </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
    set rs=server.createobject("adodb.recordset")
	sql="select * from contact where id=" & request("id") 
	rs.open sql,conn,1,3
	rs("title")=request("title")
	rs("class")=request("class")
	rs("theme")=request("theme")
	rs.update
	rs.close
	set rs=nothing	
	Response.Write "<script>alert('鎮ㄥ凡缁忔垚鍔熶慨鏀?);location='contact.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from contact where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('鍒犻櫎鎴愬姛');location='contact.asp'</script>"
end if
%>		
</body>
</html>

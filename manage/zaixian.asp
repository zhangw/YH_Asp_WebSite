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
									淇℃伅鍙戝竷&gt;&gt;淇℃伅绠＄悊</td>
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
                  <TD><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle"> 淇℃伅鐨勫鍔犮€佷慨鏀瑰拰鍒犻櫎銆?/TD>
                  <TD align="right">聽</TD>
                </TR>
                <TR>
                  <TD>聽</TD>
                  <TD align="right">聽</TD>
                </TR>
              </TBODY>
            </TABLE>
			
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="188" height="20" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">鐢宠浜?/div></th>
    <th width="88" bgcolor="#e0e9f8"  class="STYLE5">鎵€灞炲湴鍖?/th>
    <th width="161" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">搴楅潰闈㈢Н</div></th>
    <th width="87" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">鍙戝竷鏃堕棿</div></th>
    <th width="131" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">鎿嶄綔</div></th>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from zaixian " 
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
  <tr <%if RowCount mod 2=0 then%>bgColor="#fffff0"<%else%>bgcolor="#F6F6F6"<%end if%> class="unnamed1">
    <td width="188" height="20"  class="STYLE5" style="padding-left:5px;"><%=rs("name")%></td>
	<td width="88"  class="STYLE5"><%=rs("city")%></td>
	<td width="161"  class="STYLE5"><div align="center"><%=rs("mianji")%>骞虫柟绫?/div></td>
    <td  class="STYLE5"><div align="center"><%=formatdatetime(rs("submitdate"),2)%></div></td>
    <td width="131" ><div align="center"><A title="淇敼" href="?action=edit&id=<%=rs("id")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('纭疄瑕佸垹闄ゅ悧'))   href='?action=del&id=<%=rs("id")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="鍒犻櫎"></A></div></td>
  </tr>
  <%RowCount=RowCount-1
rs.movenext
loop
%>
</table>
<table width="90%" border=0 align=center cellSpacing=1 class="navi">  
  <tr>
    <td height="20" ><div align="center" class="unnamed1">绗?%= page %>椤?nbsp; <a href="?page=1<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh">棣栭〉</a> &nbsp;鍏?%=rs.PageCount%>椤?nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >涓婁竴椤?/a>
            <% else %>
        涓婁竴椤?
        <% end if %>
&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=page+1%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >涓嬩竴椤?/a>
        <% else %>
        涓嬩竴椤?
        <% end if %>
&nbsp;<select name="select" onChange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%For m = 1 To rs.PageCount%>
        <option value="?page=<%=m%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" <%if page=m then%>selected<%end if%>><%=m%></option>
        <% Next %>
      </select>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >鏈〉</a>
        <% else %>
        鏈〉
        <% end if %>
&nbsp;鎬绘暟<%= rs.recordcount %>鏉?/div></td>
  </tr>
</table>

            <%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from zaixian where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 鏌ョ湅淇℃伅</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>">
  <table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
    <tr>
      <td width="23%" height="20" align="right">鐢宠浜猴細</td>
      <td width="27%">&nbsp;<%=rs("name")%></td>
      <td width="23%" align="right">鎵€灞炲湴鍖猴細</td>
      <td width="27%">&nbsp;<%=rs("city")%></td>
    </tr>
    <tr>
      <td height="20" align="right">鐢?聽聽聽璇濓細</td>
      <td>&nbsp;<%=rs("tel")%></td>
      <td align="right">鎵嬫満锛?/td>
      <td>&nbsp;<%=rs("celltel")%></td>
    </tr>
    <tr>
      <td height="20" align="right">浼犵湡锛?/td>
      <td>&nbsp;<%=rs("fax")%></td>
      <td align="right">E-mail锛?/td>
      <td>&nbsp;<%=rs("email")%></td>
    </tr>
    <tr>
      <td height="20" align="right">鑱旂郴鍦板潃锛?/td>
      <td>&nbsp;<%=rs("adress")%></td>
      <td align="right">閭紪锛?/td>
      <td>&nbsp;<%=rs("zip")%></td>
    </tr>
    <tr>
      <td height="20" align="right">搴楅潰鍦板潃锛?/td>
      <td>&nbsp;<%=rs("dadress")%></td>
      <td align="right">搴楅潰闈㈢Н锛?/td>
      <td>&nbsp;<%=rs("mianji")%>骞虫柟绫?/td>
    </tr>
    <tr>
      <td height="20" align="right">鎷熸姇璧勭骇鍒細</td>
      <td>&nbsp;<%=rs("jibie")%></td>
      <td align="right">瑙勫垝浣曟椂寮€搴楋細</td>
      <td>&nbsp;<%=rs("heshi")%></td>
    </tr>
    <tr>
      <td height="20" align="right">鍔犵洘璁″垝锛?/td>
      <td>&nbsp;<%=rs("jihua")%></td>
      <td align="right">鎮ㄥ綋鍦颁汉鍙ｆ暟锛?/td>
      <td>&nbsp;<%=rs("renshu")%></td>
    </tr>
  </table>
</form>
<%end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from zaixian where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('鍒犻櫎鎴愬姛');location='zaixian.asp'</script>"
end if
%>		
</body>
</html>

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
<%
       Sub MainFl(celclass) 
           Dim Rs 
           Set Rs = Conn.Execute("SELECT id,class FROM productclass WHERE father=0") 
           If Not Rs.Eof Then 
               Do While Not Rs.Eof 
			       if celclass=Rs("id") then
                   response.Write "<option value=""" & Trim(Rs("id")) & """ selected>" & Trim(Rs("class")) & "</option>" 
				   else
				   response.Write "<option value=""" & Trim(Rs("id")) & """>" & Trim(Rs("class")) & "</option>" 
				   end if
                   Call Subfl(Rs("id"),Rs("class")&"鈫?,celclass) '寰幆瀛愮骇鍒嗙被 
                   response.Write "</div>" 
               Rs.MoveNext 
               If Rs.Eof Then Exit Do '闃蹭笂閫犳垚姝诲惊鐜?
               Loop 
           End If 
           Set Rs = Nothing 
       End Sub 
       '瀹氫箟瀛愮骇鍒嗙被 
       Sub SubFl(FID,StrDis,celclass) 
           Dim Rs1 
		   aa=StrDis
           Set Rs1 = Conn.Execute("SELECT id,class FROM productclass WHERE father = " & FID & "")
           If Not Rs1.Eof Then 
               Do While Not Rs1.Eof 
			       if celclass=Rs1("id") then
				   response.Write "<option value="""& Trim(Rs1("id")) & """ selected>" & StrDis & Trim(Rs1("class")) & "</option>"
				   else
                   response.Write "<option value="""& Trim(Rs1("id")) & """>" & StrDis & Trim(Rs1("class")) & "</option>" 
                   end if
				   Call SubFl(Trim(Rs1("id")),StrDis&Rs1("class")&"鈫?,celclass) '閫掑綊瀛愮骇鍒嗙被 
               Rs1.Movenext:Loop 
               If Rs1.Eof Then 
                   Rs1.Close 
                   Exit Sub 
               End If 
           End If 
           Set Rs1 = Nothing 
       End Sub
%>
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
									璁㈠崟绠＄悊&gt;&gt;璁㈠崟绠＄悊</td>
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
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="317" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">浜у搧鍚嶇О</div></th>
    <th width="126" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">璁㈣喘鏁伴噺</div></th>
    <th width="93" bgcolor="#e0e9f8"  class="STYLE5">鍗曚环</th>
    <th width="84" bgcolor="#e0e9f8"  class="STYLE5">鎬婚噾棰?/th>
    <th width="78" bgcolor="#e0e9f8"  class="STYLE5">涓嬪崟鏃ユ湡</th>
    <th width="78" bgcolor="#e0e9f8"  class="STYLE5">鐘舵€?/th>
    <th width="84" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">鎿嶄綔</div></th>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from order_list order by id desc" 
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
i=0
	do while not rs.eof and rowcount>0
	 set rs_product=server.createobject("adodb.recordset")
	sql="select * from product where id="&rs("pid") 
	rs_product.open sql,conn,1,3
%>
  <tr <%if i mod 2=0 then%>bgcolor="#fffff0"<%else%>bgcolor="#ffffff"<%end if%> class="unnamed1">
    <td width="317" height="25"  class="STYLE5" style="padding-left:5px;"><a href="../prolist.asp?id=<%=rs_product("id")%>" target="_blank"><%if len(rs_product("title"))>13 then%><%=mid(rs_product("title"),1,13)%>...<%else%><%=rs_product("title")%><%end if%></a></td>
	<td width="126"  class="STYLE5"><div align="center"><%=rs("sl")%></div></td>
    <td width="93"  class="STYLE5"><%=rs_product("price")%>鍏?/td>
    <td width="84"  class="STYLE5"><%=rs("sl")*rs_product("price")%>鍏?/td>
    <td width="78"  class="STYLE5"><%=formatdatetime(rs("submitdate"),2)%></td>
    <td width="78"  class="STYLE5"><%=rs("zt")%></td>
    <td width="84" ><div align="center"><A title="淇敼" href="?action=edit&id=<%=rs("id")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('纭疄瑕佸垹闄ゅ悧'))   href='?action=del&id=<%=rs("id")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="鍒犻櫎"></A></div></td>
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
        <a href="?page=<%=rs.pagecount%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >涓嬩竴椤?/a>
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

            <TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD>聽</TD>
                </TR>
                <TR>
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="product_add.asp">鏂板淇℃伅</A>聽 </TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from order_list where id="& request("id")
	rs.open sql,conn,2,3
	set rs_product=server.createObject("ADODB.Recordset")
	sql="select * from product where id="& rs("pid")
	rs_product.open sql,conn,2,3
	set rs_member=server.createObject("ADODB.Recordset")
	sql="select * from member where admin='"& rs("admin")&"'"
	rs_member.open sql,conn,2,3	
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><p><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 鏌ョ湅璁㈠崟淇℃伅</p>
              </td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="103"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">浜у搧鍚嶇О锛?/div></td>
      <td height="20" colspan="2" bgcolor="#FFFFFF"><%=rs_product("title")%></td>
    </tr>
        <tr>
      <td height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">鍗曚环锛?/td>
      <td colspan="2" bgcolor="#FFFFFF"><%=rs_product("price")%>鍏?/td>
    </tr>
        <tr>
          <td height="20" align="right" bgcolor="#FFFFFF" class="STYLE5">鏁伴噺锛?/td>
          <td colspan="2" bgcolor="#FFFFFF"><%=rs("sl")%></td>
        </tr>
    <tr>
      <td width="103" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">鎬婚噾棰濓細</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><%=rs_product("price")*rs("sl")%>鍏?/td>
    </tr>
	    <tr>
      <td height="20" align="right" bgcolor="#FFFFFF" class="front3">璁㈠崟鐘舵€?span class="STYLE5">锛?/span></td>
      <td width="578" height="20" bgcolor="#FFFFFF" class="front3"><input name="zt" type="text" id="zt"  value="<%=rs("zt")%>"/></td>
    </tr>

        <tr bgcolor="#A4B6D7">
          <th height="20" align="right" bgcolor="#FFFFFF">鑱旂郴浜猴細</th>
          <th height="20" align="left" bgcolor="#FFFFFF"><%=rs_member("name")%></th>
          <th width="171" height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
        </tr>
        <tr bgcolor="#A4B6D7">
          <th height="20" align="right" bgcolor="#FFFFFF">鐢佃瘽锛?/th>
          <th height="20" align="left" bgcolor="#FFFFFF"><%=rs_member("tel")%></th>
          <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
        </tr>
        <tr bgcolor="#A4B6D7">
          <th height="20" align="right" bgcolor="#FFFFFF">閭锛?/th>
          <th height="20" align="left" bgcolor="#FFFFFF"><%=rs_member("mail")%></th>
          <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
        </tr>
        <tr bgcolor="#A4B6D7">
          <th height="20" align="right" bgcolor="#FFFFFF">鍦板潃锛?/th>
          <th height="20" align="left" bgcolor="#FFFFFF"><%=rs_member("adress")%></th>
          <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
        </tr>
    <tr bgcolor="#A4B6D7">
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
      <th height="20" align="left" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="纭畾" /></th>
      <th height="20" align="center" bgcolor="#FFFFFF">&nbsp;</th>
    </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from order_list where id="&request("id")
		rs1.open sql,conn,1,3
		rs1("zt")=request("zt")
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		Response.Write "<script>alert('鎮ㄥ凡缁忔垚鍔熶慨鏀?);location='order_list.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from order_list where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('鍒犻櫎鎴愬姛');location='order_list.asp'</script>"
end if
%>		
</body>
</html>

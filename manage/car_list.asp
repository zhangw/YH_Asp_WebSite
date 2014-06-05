<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
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
								<td class="SheetSelected" vAlign="bottom" width="60">后台管理</td>
								<td valign="bottom"><img src="images/sheet_right.gif" width="25" height="18"></td>
							</tr>
						</table>
					</td>
					<td vAlign="bottom" align="right">
						<table cellSpacing="2" cellPadding="0" width="99%" border="0">
							<tr>
								<td align="right">您的位置：<A href="index.asp" target="_top">后台管理</A> &gt;&gt; 
									产品管理&gt;&gt;产品管理</td>
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
                  <TD height="21"><IMG height="20" src="images/icon_info.gif" width="20" align="absMiddle">产品 信息的增加、修改和删除。</TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
                <TR>
                  <TD></TD>
                  <TD align="right">&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
			
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#e0e9f8" class="unnamed1">
    <th width="273" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">标题</div></th>
    <th width="142" bgcolor="#e0e9f8"  class="STYLE5">交易地区</th>
    <th width="118" bgcolor="#e0e9f8"  class="STYLE5">行驶里程</th>
    <th width="171" bgcolor="#e0e9f8"  class="STYLE5">发布人</th>
    <th width="101" bgcolor="#e0e9f8"  class="STYLE5">热门车源</th>
    <th width="146" bgcolor="#e0e9f8"  class="STYLE5">发布时间</th>
    <th width="94" height="25" bgcolor="#e0e9f8"  class="STYLE5"><div align="center">操作</div></th>
  </tr>
  <%     set rs=server.createobject("adodb.recordset")
	sql="select * from car where sh=1 " 
	sql=sql & " order by id desc"
	rs.open sql,conn,1,3
if  rs.eof then
 response.Write "没有记录!"
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
  <tr <%if rowcount mod 2=0 then%>bgcolor="#fffff0"<%else%>bgcolor="#ffffff"<%end if%> class="unnamed1">
    <td height="20"  class="STYLE5" style="padding-left:5px;"><%if len(rs("title"))>20 then%><%=mid(rs("title"),1,20)%>...<%else%><%=rs("title")%><%end if%></td>
	
	<td width="142" align="center"  class="STYLE5"><div align="center"><%=rs("dqu")%></div></td>

	<td width="118" align="center"  class="STYLE5"><%=rs("xslc")%></td>
<%
set rs1=server.createobject("adodb.recordset")
sql="select * from [user] where admin='"&rs("admin")&"'" 
rs1.open sql,conn,1,3
%>
	<td width="171" align="center"  class="STYLE5"><nobr><%=rs1("company")%></nobr></td>

	<td width="101" align="center"  class="STYLE5"><%if rs("rm")=1 then%>
      <a href="?action=tuij&id=<%=rs("id")%>&page=<%=int(request("page"))%>">取消推荐</a>
      <%else%>
      <a href="?action=bt&id=<%=rs("id")%>&page=<%=int(request("page"))%>">推荐</a>
    <%end if%></td>
	<td width="146" align="center"  class="STYLE5"><%=rs("time")%></td>
    <td width="94" ><div align="center"><A title="修改" href="?action=edit&id=<%=rs("id")%>"><IMG height="20" src="images/icon_edit.gif" width="20" border="0"></A><a   href="#"onclick="javascript:if   (confirm('确实要删除吗'))   href='?action=del&id=<%=rs("id")%>';else   return;"><IMG src="images/icon_del.gif" alt="" width="20" height="20" border="0" title="删除"></A></div></td>
  </tr>
  <%RowCount=RowCount-1
rs.movenext
loop
%>
</table>
<table width="90%" border=0 align=center cellSpacing=1 class="navi">  
  <tr>
    <td height="20" ><div align="center" class="unnamed1">第<%= page %>页&nbsp; <a href="?page=1<%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh">首页</a> &nbsp;共<%=rs.PageCount%>页&nbsp;
            <% if page>1 then %>
            <a href="?page=<%= page-1 %><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >上一页</a>
            <% else %>
        上一页
        <% end if %>
&nbsp;<span class="A3"> </span>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >下一页</a>
        <% else %>
        下一页
        <% end if %>
&nbsp;<select name="select" onChange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%For m = 1 To rs.PageCount%>
        <option value="?page=<%=m%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" <%if page=m then%>selected<%end if%>><%=m%></option>
        <% Next %>
      </select>
        <% if page<rs.pagecount then %>
        <a href="?page=<%=rs.pagecount%><%if request("danwei")<>"" then%>&danwei=<%=request("danwei")%><%end if%><%if request("name")<>"" then%>&name=<%=request("name")%><%end if%>" class="hh" >末页</a>
        <% else %>
        末页
        <% end if %>
&nbsp;总数<%= rs.recordcount %>条</div></td>
  </tr>
</table>

            <TABLE width="90%" border="0" align="center" cellPadding="0" cellSpacing="0">
              <TBODY>
                <TR>
                  <TD></TD>
                </TR>
                <TR>
                  <TD><IMG height="20" src="images/icon_new.gif" width="20" align="absMiddle"> <A href="product_add.asp">新增信息</A></TD>
                </TR>
              </TBODY>
</TABLE>
<%end if%>	
<%
if request("action")="edit" then
	set rs=server.createObject("ADODB.Recordset")
	sql="select * from car where id="& request("id")
	rs.open sql,conn,2,3
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 修改产品信息</td>
            </tr>
            <tr>
              <td><hr width="100%" color="#000000" noshade="noshade" size="2" />
              </td>
            </tr>
          </tbody>
</table>
<script type="text/javascript">
	function check(){
		if(document.formm.title.value == ""){
			alert("请输入产品的名称");
			return false;
		}
		if(document.formm.admin.value == ""){
			alert("请选择发布人");
			return false;
		}

		return true;
	}
</script>

<script language=javascript> 
function OnClick(){ 
S=window.showModalDialog("yonghu.asp"); 
document.all.item("admin").value=S; 
} 
</script>
<form name="formm" method="post" action="?action=editsave&id=<%=request("id")%>" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="126"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">信息标题：</div></td>
      <td width="934" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;" value="<%=rs("title")%>"></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">车子价格：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="price" type="text" id="price" style="width:300px;" value="<%=rs("price")%>"/></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">交易地区：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="dqu" type="text" id="dqu" style="width:300px;" value="<%=rs("dqu")%>"/></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">上牌时间：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="spsj" type="text" id="spsj" style="width:300px;" value="<%=rs("spsj")%>"/></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">行驶里程：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="xslc" type="text" id="xslc" style="width:300px;" value="<%=rs("xslc")%>"/></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">有效时间：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="yxq" type="text" id="yxq" style="width:300px;" value="<%=rs("yxq")%>"/></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">车辆颜色：</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="clys" type="text" id="clys" style="width:300px;" value="<%=rs("clys")%>"/></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">上牌地区：</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="spdq" type="text" id="spdq" style="width:300px;" value="<%=rs("spdq")%>"/></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">车辆类别：</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="cllb" id="cllb">
        <option value="轿车" <%if rs("cllb")="轿车" then%> selected="selected"<%end if%>>轿车</option>
        <option value="跑车" <%if rs("cllb")="跑车" then%> selected="selected"<%end if%>>跑车</option>
        <option value="SUV越野车" <%if rs("cllb")="SUV越野车" then%> selected="selected"<%end if%>>SUV越野车</option>
        <option value="MPV" <%if rs("cllb")="MPV" then%> selected="selected"<%end if%>>MPV</option>
        <option value="货车" <%if rs("cllb")="货车" then%> selected="selected"<%end if%>>货车</option>
        <option value="皮卡" <%if rs("cllb")="皮卡" then%> selected="selected"<%end if%>>皮卡</option>
        <option value="客车" <%if rs("cllb")="客车" then%> selected="selected"<%end if%>>客车</option>
        <option value="面包车" <%if rs("cllb")="面包车" then%> selected="selected"<%end if%>>面包车</option>
        <option value="其他" <%if rs("cllb")="其他" then%> selected="selected"<%end if%>>其他</option>
      </select>
      </td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">车辆用途：</td>
      <td colspan="2" bgcolor="#FFFFFF"><label>
        <select name="clyt" id="clyt">
          <option value="营运" <%if rs("clyt")="营运" then%> selected="selected"<%end if%>>营运</option>
          <option value="非营运" <%if rs("clyt")="非营运" then%> selected="selected"<%end if%>>非营运</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">公车/私车：</td>
      <td colspan="2" bgcolor="#FFFFFF"><input type="radio" name="gcsc" value="公车" <%if rs("gcsc")="公车" then%>checked="checked"<%end if%>/>
        公车
          <input name="gcsc" type="radio" value="私车" <%if rs("gcsc")="私车" then%>checked="checked"<%end if%> />
      私车</td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">变 速 器：</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="bsq" id="bsq">
        <option value="自动" <%if rs("bsq")="自动" then%>selected="selected"<%end if%>>自动</option>
        <option value="手动" <%if rs("bsq")="手动" then%>selected="selected"<%end if%>>手动</option>
        <option value="手自一体" <%if rs("bsq")="手自一体" then%>selected="selected"<%end if%>>手自一体</option>
        <option value="CVT" <%if rs("bsq")="CVT" then%>selected="selected"<%end if%>>CVT</option>
      </select>
      </td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">保险情况：</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="bx" type="text" id="bx" style="width:300px;" <%=rs("bx")%>/></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">价格范围：</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="jgfw" id="jgfw">
        <option value="3万以下" <%if rs("jgfw")="3万以下" then%>selected="selected"<%end if%>>3万以下</option>
        <option value="3-6万" <%if rs("jgfw")="3-6万" then%>selected="selected"<%end if%>>3-6万</option>
        <option value="6-10万" <%if rs("jgfw")="6-10万" then%>selected="selected"<%end if%>>6-10万</option>
        <option value="10-15万" <%if rs("jgfw")="10-15万" then%>selected="selected"<%end if%>>10-15万</option>
        <option value="15-20万" <%if rs("jgfw")="15-20万" then%>selected="selected"<%end if%>>15-20万</option>
        <option value="20-30万" <%if rs("jgfw")="20-30万" then%>selected="selected"<%end if%>>20-30万</option>
        <option value="30万以上" <%if rs("jgfw")="30万以上" then%>selected="selected"<%end if%>>30万以上</option>
      </select>
      </td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">使用时间：</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="sjfw" id="sjfw">
        <option value="1年以下" <%if rs("sjfw")="1年以下" then%>selected="selected"<%end if%>>1年以下</option>
        <option value="1-3年" <%if rs("sjfw")="1-3年" then%>selected="selected"<%end if%>>1-3年</option>
        <option value="3-5年" <%if rs("sjfw")="3-5年" then%>selected="selected"<%end if%>>3-5年</option>
        <option value="5-7年" <%if rs("sjfw")="5-7年" then%>selected="selected"<%end if%>>5-7年</option>
        <option value="7-10年" <%if rs("sjfw")="7-10年" then%>selected="selected"<%end if%>>7-10年</option>
        <option value="10年以上" <%if rs("sjfw")="10年以上" then%>selected="selected"<%end if%>>10年以上</option>
      </select>
      </td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">详细情况：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"><%=rs("content")%></textarea>
	                <iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe></td>
    </tr>
	    
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">汽车照片1<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <input name="image" type="hidden" id="image" value="<%=rs("pic1")%>"/><iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe> </td>
    </tr>			
	   
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">汽车照片2：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"> <input name="image2" type="hidden" id="image2" value="<%=rs("pic2")%>"/><iframe id="1" src="upfile1.asp?path=product&name=image2" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">汽车照片3：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"> <input name="image3" type="hidden" id="image3" value="<%=rs("pic3")%>"/><iframe id="1" src="upfile1.asp?path=product&name=image3" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">汽车照片4：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"> <input name="image4" type="hidden" id="image4" value="<%=rs("pic4")%>"/><iframe id="1" src="upfile1.asp?path=product&name=image4" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
    </tr>

    <tr>
      <td width="126" align="right" bgcolor="#FCFCFC" class="front3"><div align="right">发布企业:</div></td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="admin" type="text" id="admin" value="<%=rs("admin")%>"/>
      &nbsp;<a href="#" class="link1" onClick="OnClick()">获取发布人</a></td>
    </tr>
    <tr bgcolor="#A4B6D7">
      <td width="126" height="20" align="right" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <div align="right"></div></td>
      <td height="20" align="left" bgcolor="#FFFFFF">
	  <input type="submit" name="Submit" value="确定" />	  </td>
    </tr>
  </table>
</form>
<%end if
if request("action")="editsave" then
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from car where id="&request("id")
		rs1.open sql,conn,1,3
		rs1("title")=request("title")'标题
		rs1("content")=request("theme")'内容
		rs1("price")=request("price")'产品类别
		rs1("dqu")=request("dqu")'产品类别
		rs1("spsj")=request("spsj")'产品的存放文件名
		rs1("xslc")=request("xslc")'产品的存放文件
		rs1("yxq")=request("yxq")'产品的存放文件
		rs1("clys")=request("clys")'产品的存放文件
		rs1("spdq")=request("spdq")'产品的存放文件
		rs1("cllb")=request("cllb")'产品的存放文件
			rs1("clyt")=request("clyt")'产品的存放文件
			rs1("gcsc")=request("gcsc")'产品的存放文件
			rs1("bsq")=request("bsq")'产品的存放文件
			rs1("bx")=request("bx")'产品的存放文件
			rs1("pic1")=request("image")'产品的存放文件
			rs1("pic2")=request("image2")'产品的存放文件
			rs1("pic3")=request("image3")'产品的存放文件
			rs1("pic4")=request("image4")'产品的存放文件
			rs1("admin")=request("admin")'产品的存放文件	
						rs1("rm")=int(request("rm"))'产品的存放文件	
			rs1("yd")=int(request("yd"))'产品的存放文件	
		rs1("jgfw")=request("jgfw")'产品的存放文件			
		rs1("sjfw")=request("sjfw")'产品的存放文件			
		rs1("time")=now()
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
		Response.Write "<script>alert('您已经成功修改');location='car_list.asp'</script>"
end if
if request("action")="del" then
		set rs=server.createObject("ADODB.Recordset")
		sql="select * from car where id="& request("id") 
		rs.open sql,conn,2,3
		if not rs.eof then
		rs.delete
		rs.update
		rs.requery
		end if
		rs.close
		set rs=nothing
		Response.Write "<script>alert('删除成功');location='car_list.asp'</script>"
end if
if request("action")="tuij" then
    theme=trim(request("theme"))
    set rs=server.createobject("adodb.recordset")
	sql="select * from car where id="&request("id") 
	rs.open sql,conn,1,3
	rs("rm")=0
	rs.update
	rs.close
	set rs=nothing	
	Response.Redirect "car_list.asp?page="&request("page")
end if
if request("action")="bt" then
    theme=trim(request("theme"))
    set rs=server.createobject("adodb.recordset")
	sql="select * from car where id="&request("id") 
	rs.open sql,conn,1,3
	rs("rm")=1
	rs.update
	rs.close
	set rs=nothing	
	Response.Redirect "car_list.asp?page="&request("page")
end if
%>		
</body>
</html>

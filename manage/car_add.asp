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
									二手车信息&gt;&gt;发布信息</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgColor="#cccccc" colSpan="2"><img height="1" alt="" src="" width="1" name=""></td>
				</tr>
			</table><br>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td valign="bottom" class="SubTitle"><img height="20" src="images/icon_info.gif" width="20" align="absmiddle" /> 新增信息</td>
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
			alert("请输入标题");
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
<form name="formm" method="post" action="?action=aa" onsubmit="return check()">
<table width="90%" border=0 align=center cellSpacing=1 class="navi">
    <tr>
      <td width="126"  height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">信息标题：</div></td>
      <td width="934" height="20" colspan="2" bgcolor="#FFFFFF"><input name="title" type="text" id="title" style="width:300px;"></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="STYLE5"><div align="right">车子价格：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="price" type="text" id="price" style="width:300px;" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">交易地区：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="dqu" type="text" id="dqu" style="width:300px;" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">上牌时间：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="spsj" type="text" id="spsj" style="width:300px;" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">行驶里程：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="xslc" type="text" id="xslc" style="width:300px;" /></td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">有效时间：</div></td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="yxq" type="text" id="yxq" style="width:300px;" /></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">车辆颜色：</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="clys" type="text" id="clys" style="width:300px;" /></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">上牌地区：</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="spdq" type="text" id="spdq" style="width:300px;" /></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">车辆类别：</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="cllb" id="cllb">
        <option value="轿车">轿车</option>
        <option value="跑车">跑车</option>
        <option value="SUV越野车">SUV越野车</option>
        <option value="MPV">MPV</option>
        <option value="货车">货车</option>
        <option value="皮卡">皮卡</option>
        <option value="客车">客车</option>
        <option value="面包车">面包车</option>
        <option value="其他">其他</option>
      </select>      </td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">车辆用途：</td>
      <td colspan="2" bgcolor="#FFFFFF"><label>
        <select name="clyt" id="clyt">
          <option value="营运">营运</option>
          <option value="非营运" selected="selected">非营运</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">公车/私车：</td>
      <td colspan="2" bgcolor="#FFFFFF"><input type="radio" name="gcsc" value="公车" />
        公车
          <input name="gcsc" type="radio" value="私车" checked="checked" />
      私车</td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">变 速 器：</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="bsq" id="bsq">
        <option value="自动" selected="selected">自动</option>
        <option value="手动">手动</option>
        <option value="手自一体">手自一体</option>
        <option value="CVT">CVT</option>
      </select>      </td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">保险情况：</td>
      <td colspan="2" bgcolor="#FFFFFF"><input name="bx" type="text" id="bx" style="width:300px;" /></td>
    </tr>

    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">价格范围：</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="jgfw" id="jgfw">
        <option value="3万以下" selected="selected">3万以下</option>
        <option value="3-6万">3-6万</option>
        <option value="6-10万">6-10万</option>
        <option value="10-15万">10-15万</option>
        <option value="15-20万">15-20万</option>
        <option value="20-30万">20-30万</option>
        <option value="30万以上">30万以上</option>
      </select>
      </td>
    </tr>
    <tr>
      <td height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5">使用时间：</td>
      <td colspan="2" bgcolor="#FFFFFF"><select name="sjfw" id="sjfw">
        <option value="1年以下" selected="selected">1年以下</option>
        <option value="1-3年">1-3年</option>
        <option value="3-5年">3-5年</option>
        <option value="5-7年">5-7年</option>
        <option value="7-10年">7-10年</option>
        <option value="10年以上">10年以上</option>
      </select>
      </td>
    </tr>
    <tr>
      <td width="126" height="20" align="right" valign="top" bgcolor="#FFFFFF" class="STYLE5"><div align="right">详细情况：</div></td>
      <td colspan="2" bgcolor="#FFFFFF">
	                <textarea name="theme" style="display:none"></textarea>
	                <iframe id="editor2" src="../Editor/eWebEditor.asp?id=theme" frameborder=1 scrolling=no width="550" height="405"></iframe></td>
    </tr>
	    
    <tr>
      <td width="126" height="20" align="right" bgcolor="#FFFFFF" class="front3"><div align="right">汽车照片1<span class="STYLE5">：</span></div></td>
      <td height="20" bgcolor="#FFFFFF" class="front3">
      <input name="image" type="hidden" id="image"/><iframe id="1" src="upfile1.asp?path=product&name=image" frameborder="0" scrolling="No" width="300" height="25"></iframe> </td>
    </tr>			
	   
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">汽车照片2：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"> <input name="image2" type="hidden" id="image2"/><iframe id="1" src="upfile1.asp?path=product&name=image2" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">汽车照片3：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"> <input name="image3" type="hidden" id="image3"/><iframe id="1" src="upfile1.asp?path=product&name=image3" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FCFCFC" class="front3">汽车照片4：</td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"> <input name="image4" type="hidden" id="image4"/><iframe id="1" src="upfile1.asp?path=product&name=image4" frameborder="0" scrolling="No" width="300" height="25"></iframe></td>
    </tr>

    <tr>
      <td width="126" align="right" bgcolor="#FCFCFC" class="front3"><div align="right">发布企业:</div></td>
      <td height="25" colspan="2" bgcolor="#FCFCFC" class="front3"><input name="admin" type="text" id="admin"/>
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
<%
if request("action")="aa" then

		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from car "
		rs1.open sql,conn,1,3
		rs1.addnew
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
		rs1("sjfw")=request("sjfw")'产品的存放文件
		rs1("jgfw")=request("jgfw")'产品的存放文件
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
			rs1("sh")=1
		rs1("time")=now()
		rs1.update
		rs1.requery
		rs1.close
		set rs1=nothing
	Response.Write "<script>alert('您已经成功添加');location='car_add.asp'</script>"
end if	
%>
</body>
</html>


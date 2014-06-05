<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>乔恩传媒</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script language=javascript> 
function OnClick(aa){ 
window.returnValue=aa; 
window.close(); 
} 
function OnClick1(){ 
document.formm.submit(); 
}
</script>
</head>

<body>
<br />
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
     <form id="formm" name="formm" method="post"  onsubmit="return OnClick1('dssdf')"><td>
    <iframe id="1" src="upfile1.asp?name=<%=request("name")%>" frameborder="0" scrolling="No" width="300" height="25"></iframe>
   
    
    <input type="text" name="<%=request("name")%>" onchange="OnClick()" value="sdfds"/></td></form>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="file:///C|/Documents and Settings/Administrator/桌面/11.png" width="292" height="326"  name="pic" id="pic"/></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>

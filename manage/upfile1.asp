<html>
<head>

<title>图片上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body leftmargin="0" topmargin="0" >

<script type="text/javascript">
	function check(){
		if(document.form1.file1.value == ""){
			alert("请选择上传文件!");
			return false;
		}else{
			return true;
		}
	}
</script>
<form action="upfile2.asp?name=<%=request("name")%>" method="post" enctype="multipart/form-data" name="form1" onSubmit="return check()">
  <input type="file" name="file1" value="">
  <%session("name")=request("name")%>
  <input type="submit" name="Submit" value="提交">
</form>
</body>
</html>      



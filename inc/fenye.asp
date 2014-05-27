<script type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>

<script>
function go(){
	var v=document.getElementById("sel").value;
	window.location.href=v;
	}
</script>
<%
 function movepage(all,page,max,url,pagesizes)

	if page = 1 or page = "" then 
		response.write "&nbsp;首页 |"
		response.write "&nbsp;上一页 |"
	else 
		response.write "&nbsp;<a href='"&url&"page=1' class='css1' title='返回到第一页'>首页</a> | "
		response.write "&nbsp;<a href='"&url&"page="&page-1&"' class='css1' title='返回到上一页'>上一页</a> |"
	end if 
	if page >= max then 
		response.write "&nbsp;下一页 |"
		response.write "&nbsp;尾页 "
	else 
		response.write "&nbsp;<a href='"&url&"page="&page+1&"' class='css1' title='跳转到下一页'>下一页</a> |"
		response.write "&nbsp;<a href='"&url&"page="&max&"' class='css1' title='跳转到最后一页'>尾页</a> "
	end if 
	response.write "&nbsp;转到第"
	response.write "<select id='sel' name='sel' onchange=MM_jumpMenu('parent',this,0)>"
	for x=1 to max 
		if x = page then 
			response.write"<option value='"&url&"page="&x&"' selected>"&x&"</option>"
		else
			response.write"<option value='"&url&"page="&x&"'>"&x&"</option>"
		end if
	next 
        response.write "</select>页"
end function
%>
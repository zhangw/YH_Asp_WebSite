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
		response.write "&nbsp;��ҳ |"
		response.write "&nbsp;��һҳ |"
	else 
		response.write "&nbsp;<a href='"&url&"page=1' class='css1' title='���ص���һҳ'>��ҳ</a> | "
		response.write "&nbsp;<a href='"&url&"page="&page-1&"' class='css1' title='���ص���һҳ'>��һҳ</a> |"
	end if 
	if page >= max then 
		response.write "&nbsp;��һҳ |"
		response.write "&nbsp;βҳ "
	else 
		response.write "&nbsp;<a href='"&url&"page="&page+1&"' class='css1' title='��ת����һҳ'>��һҳ</a> |"
		response.write "&nbsp;<a href='"&url&"page="&max&"' class='css1' title='��ת�����һҳ'>βҳ</a> "
	end if 
	response.write "&nbsp;ת����"
	response.write "<select id='sel' name='sel' onchange=MM_jumpMenu('parent',this,0)>"
	for x=1 to max 
		if x = page then 
			response.write"<option value='"&url&"page="&x&"' selected>"&x&"</option>"
		else
			response.write"<option value='"&url&"page="&x&"'>"&x&"</option>"
		end if
	next 
        response.write "</select>ҳ"
end function
%>
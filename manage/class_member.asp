<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
	color: #000000;
}
a {
	font-size: 12px;
	color: #000000;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #000000;
}
a:hover {
	text-decoration: none;
	color: #000000;
}
a:active {
	text-decoration: none;
	color: #000000;
}
-->
</style>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
.style31 {color: #20419C}
.style2 {font-size: 9pt}
-->
</style>
<%set rsx=server.createobject("adodb.recordset")
sql = "select * from smallclass order by ID asc"
rsx.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rsx.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rsx("smallclass"))%>",<%= rsx("bigclass")%>,"<%= trim(rsx("id"))%>");
        <%
        count = count + 1
        rsx.movenext
        loop
        rsx.close
        %>
onecount=<%=count%>;
   function changelocation(locationid)
    {
	var m=0;
    document.formm.SmallClassName.length = 0; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
			 if (m==0)
			{
			m=subcat[i][2];
			}
                document.formm.SmallClassName.options[document.formm.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    
</script>

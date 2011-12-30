<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Installment Due Dates</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Save(){
		var temp = "";	
		if (document.frmpop6.count.value=="1") {
			temp = document.frmpop6.DueDate.value;
		} else {
			temp = document.frmpop6.DueDate[0].value;
			for (var i = 1; i < document.frmpop6.count.value; i++){
				temp = temp + ":" + document.frmpop6.DueDate[i].value;
			}			
		}
		window.returnValue = temp;
		window.close();
	}	
	</script>
</head>
<body>
<form name="frmpop6">
<h5>Installment Due Dates</h5>
<hr>
<table cellpadding="1" cellspacing="1" align="center">
<% 
for (var i=0; i < Request.QueryString("num"); i++) {
%>
    <tr> 
		<td nowrap>Installment #<%=i+1%> Due Date</td>
		<td nowrap><input type="text" name="DueDate" size="11" maxlength="10" onChange="FormatDate(this)"></td>
    </tr>
<%
}
%>
</table>
<br><br>
<input type="button" value="Save" onClick="Save();" class="btnstyle">
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
<input type="hidden" name="count" value="<%=Request.QueryString("num")%>">
</form>
</body>
</html>
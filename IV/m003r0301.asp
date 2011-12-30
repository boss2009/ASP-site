<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_school2(0,'',0,0,0,0,0,1,0,'',2,'Q',0)}";
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();
%>
<html>
<head>
	<title>Inventory-Clients by Institution</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function ViewReport(Type) {
		switch (Type) {
			case "Report" : 
				document.frm0301.action = "m003r0301q.asp?insSchool_id="+document.frm0301.Institution.value;
				//alert(document.frm0301.action);
				document.frm0301.submit();
			break;
			case "Excel" : 
				var objNewWin ;
				objNewWin  = window.open('m003r0301excel.asp?insSchool_id='+document.frm0301.Institution.value,'w03r0301x',config='height=380,width=680,resizable=1,status=1,menubar=1');
			break;
		}         
	}
	</Script>
</head>
<body onLoad="document.frm0301.Institution.focus();">
<form name="frm0301" METHOD="POST">
<h5>Inventory Report</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap colspan="2">This report returns all inventories whose current user is attending selected institution.</td>
	</tr>
	<tr> 
		<td nowrap>Institution:</td>
		<td nowrap><select name="Institution" tabindex="1" accesskey="F">
		<% 
		while (!rsInstitution.EOF) {
		%>
			<option value="<%=(rsInstitution.Fields.Item("insSchool_id").Value)%>" <%=((rsInstitution.Fields.Item("insSchool_id").Value == 0)?"SELECTED":"")%> ><%=(rsInstitution.Fields.Item("chvSchool_Name").Value)%></option>
		<%
			rsInstitution.MoveNext();
		}
		%>
		</select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Submit" onClick="return ViewReport('Report')" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Excel" onClick="return ViewReport('Excel')" tabindex="5" class="btnstyle"></td>
    </tr>
</table>
</form>
</body>
</html>
<%
rsInstitution.Close();
%>
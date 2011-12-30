<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup(714)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

var rsRelationship = Server.CreateObject("ADODB.Recordset");
rsRelationship.ActiveConnection = MM_cnnASP02_STRING;
rsRelationship.Source = "{call dbo.cp_ASP_Lkup(14)}";
rsRelationship.CursorType = 0;
rsRelationship.CursorLocation = 2;
rsRelationship.LockType = 3;
rsRelationship.Open();
%>
<html>
<head>
	<title>Contact Relationship Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function ViewReport(Type) {
		var MM_filter = "insRtnship_id=" + document.frm0402.ContactRelationship.value;
		document.frm0402.MM_param.value = MM_filter; 		
		switch (Type) {
			case "Report" : 
				document.frm0402.action = "m001r0401q.asp?flag=2" ;
			break;
			case "Excel" : 
				var objNewWin ;
				objNewWin  = window.open('m001r0401excel.asp?flag=2&insRtnship_id='+document.frm0402.ContactRelationship.value,'w01r0401x',config='height=380,width=680,resizable=1,status=1,menubar=1');
				//objNewWin.blur();
			break;
		}         
	}
	</Script>
</head>
<body onLoad="document.frm0402.ContactRelationship.focus();">
<form name="frm0402" action="" METHOD="GET">
<h5>Contact Relationship Report</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td colspan="2">This report returns clients who have matching contact relationship.</td>
	</tr>
	<tr> 
		<td>Contact Relationship:</td>
		<td><select name="ContactRelationship" tabindex="1" accesskey="F">
			<% 
			while (!rsRelationship.EOF) {
			%>
				<option value="<%=(rsRelationship.Fields.Item("insRtnship_id").Value)%>" <%=((rsRelationship.Fields.Item("insRtnship_id").Value == 0)?"SELECTED":"")%> ><%=(rsRelationship.Fields.Item("chvname").Value)%></option>
			<%
				rsRelationship.MoveNext();
			}
			%>
		</select></td>
	</tr>
    <tr> 
		<td>Sort By Column:</td>
		<td>
			<select name="SortByColumn" tabindex="2">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
			<select name="OrderBy" tabindex="3" accesskey="L">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Submit" onClick="return ViewReport('Report')" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Excel" onClick="return ViewReport('Excel')" tabindex="5" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_flag" value="true">
<input type="hidden" name="MM_param" value="">
</form>
</body>
</html>
<%
rsCol.Close();
rsRelationship.Close();
%>
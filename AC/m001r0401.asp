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
%>
<html>
<head>
	<title>Client Contact Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				ViewReport('Report');
				document.frm01r0401.submit();
			break;		
		   	case 69 :
				//alert("E");
				ViewReport('Excel')
			break;
		}
	}
	</script>	
	<Script language="Javascript">
	if (window.focus) self.focus();
	function ViewReport(Type) {
		document.frm01r0401.MM_param.value = " chvContactName  = '" + document.frm01r0401.ContactName.value + "' ";		
		switch (Type) {
			case "Excel" : 
				var objNewWin ;
				wdest = "m001r0401excel.asp?flag=1&SortByColumn=" + document.frm01r0401.SortByColumn.value + "&OrderBy=" + document.frm01r0401.OrderBy.value + "&MM_param="+document.frm01r0401.MM_param.value ;
				objNewWin  = window.open(wdest,'w01r0401x',config='height=380,width=680,resizable=1,status=1,menubar=1');
				//objNewWin.blur();
			break;
			case "Report":
				document.frm01r0401.action="m001r0401q.asp?flag=1";
			break;
		}     
	}
	</Script>
</head>
<body onLoad="document.frm01r0401.ContactName.focus();">
<form name="frm01r0401" METHOD="GET">
<h5>Contact Report</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap colspan="2">This report returns a list of clients who have the following as their contact.</td>
	</tr>
    <tr> 
		<td nowrap>Contact Name:</td>
		<td nowrap><input type="text" name="ContactName" maxlength="100"></td>
    </tr>
    <tr> 
		<td nowrap>Sort By Column:</td>
		<td nowrap>
			<select name="SortByColumn">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvName").Value)%>
			<%
				rsCol.MoveNext();
			}
			%>
	        </select>
			<select name="OrderBy">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
	        </select>		
		</td>
	</tr>
</table>
<hr>
<table>
    <tr> 
		<td><input type="button" value="Submit" onClick="return ViewReport('Report')" class="btnstyle"></td>
		<td><input type="button" value="Excel Export" onClick="return ViewReport('Excel')" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_flag"  value="true">
<input type="hidden" name="MM_param" value="">
</form>
</body>
</html>
<%
rsCol.Close();
%>
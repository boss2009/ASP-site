<!--------------------------------------------------------------------------
* File Name: m014r0102.asp
* Title: Delivery Performance Report
* Main SP: cp_PR_Pfrm_Rpt
* Description: Delivery Performance Report.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsVendor = Server.CreateObject("ADODB.Recordset");
rsVendor.ActiveConnection = MM_cnnASP02_STRING;
rsVendor.Source = "{call dbo.cp_ASP_Lkup(3)}";
rsVendor.CursorType = 0;
rsVendor.CursorLocation = 2;
rsVendor.LockType = 3;
rsVendor.Open();
%>
<html>
<head>
	<title>Delivery Performance Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Search(output){
		if ((!CheckDate(document.frm0102.StartDate.value)) || (document.frm0102.StartDate.value=="")) {
			alert("Invalid Starting Date.");
			document.frm0102.StartDate.focus();
			return ;
		}
		if ((!CheckDate(document.frm0102.EndDate.value)) || (document.frm0102.EndDate.value=="")) {
			alert("Invalid End Date.");
			document.frm0102.EndDate.focus();
			return ;
		}
		if (!CheckDateBetween(document.frm0102.StartDate.value+" and "+document.frm0102.EndDate.value)) return ;
		if (output=="excel") {
			var ExcelSearch = window.open("m014r0102excel.asp?Vendor="+document.frm0102.Vendor.value+"&VendorName="+document.frm0102.Vendor.options[document.frm0102.Vendor.selectedIndex].text+"&StartDate="+document.frm0102.StartDate.value+"&EndDate="+document.frm0102.EndDate.value);
		} else {
			document.frm0102.action = "m014r0102q.asp";		
			document.frm0102.VendorName.value=document.frm0102.Vendor.options[document.frm0102.Vendor.selectedIndex].text;
			document.frm0102.submit();
		}
	}
	</script>	
</head>
<body onload="document.frm0102.Vendor.focus();">
<form name="frm0102" method="POST">
<h3>Delivery Performance - Vendor:</h3>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Vendor:</td>
		<td><select name="Vendor" tabindex="1" accesskey="L">
			<% 
			while (!rsVendor.EOF) {
			%>
				<option value="<%=rsVendor.Fields.Item("intCompany_id").Value%>"><%=rsVendor.Fields.Item("chvCompany_Name").Value%>
			<% 
				rsVendor.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td>Start Date:</td>
		<td><input type="text" name="StartDate" tabindex="2" maxlength="10" size="11" onChange="FormatDate(this)"></td>
	</tr>
	<tr>
		<td>End Date:</td>
		<td><input type="text" name="EndDate" tabindex="3" maxlength="10" size="11" accesskey="L" onChange="FormatDate(this)"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Generate Report" onClick="Search('asp');" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Excel" onClick="Search('excel');" tabindex="5" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="VendorName">
</form>
</body>
</html>
<%
rsVendor.Close();
%>
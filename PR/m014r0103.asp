<!--------------------------------------------------------------------------
* File Name: m014r0103.asp
* Title: Delivery Performance Criteria
* Main SP: 
* Description: Delivery Performance Report.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<html>
<head>
	<title>Delivery Performance Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	   
	
	function Toggle(){
		openWindow("m014p01FSq.asp","");	
	}
	
	function Search(output){
		if ((!CheckDate(document.frm14s01.StartDate.value)) || (document.frm14s01.StartDate.value=="")) {
			alert("Invalid Starting Date.");
			document.frm14s01.StartDate.focus();
			return ;
		}
		if ((!CheckDate(document.frm14s01.EndDate.value)) || (document.frm14s01.EndDate.value=="")) {
			alert("Invalid End Date.");
			document.frm14s01.EndDate.focus();
			return ;
		}
		if (!CheckDateBetween(document.frm14s01.StartDate.value+" and "+document.frm14s01.EndDate.value)) return ;
		if (output=="excel") {
			var ExcelSearch = window.open("m014r0103excel.asp?ClassSearchID="+document.frm14s01.ClassSearchID.value+"&ClassSearchText="+document.frm14s01.ClassSearchText.value+"&StartDate="+document.frm14s01.StartDate.value+"&EndDate="+document.frm14s01.EndDate.value);
		} else {
			document.frm14s01.action = "m014r0103q.asp";		
			document.frm14s01.submit();
		}
	}
	</script>	
</head>
<body onLoad="document.frm14s01.List.focus();">
<form name="frm14s01" method="POST">
<h3>Delivery Performance - Vendor:</h3>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Class:</td>
		<td nowrap>
			<input type="text" name="ClassSearchText" tabindex="1" accesskey="L" size="40" readonly>
			<input type="button" name="List" value="List" onClick="Toggle();" tabindex="2" class="btnstyle">			
		</td>
	</tr>
	<tr>
		<td nowrap>Start Date:</td>
		<td nowrap><input type="text" name="StartDate" tabindex="3" maxlength="10" size="11" onChange="FormatDate(this)"></td>
	</tr>
	<tr>
		<td nowrap>End Date:</td>
		<td nowrap><input type="text" name="EndDate" tabindex="4" maxlength="10" size="11" accesskey="L" onChange="FormatDate(this)"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Generate Report" onClick="Search('asp');" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Excel" onClick="Search('excel');" tabindex="6" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ClassSearchID">
</form>
</body>
</html>
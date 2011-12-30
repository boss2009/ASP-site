<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_Get_staff_loan("+Request.QueryString("insStaff_id")+",0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();
%>
<html>
<head>
	<title>Loan History</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</Script>
</head>
<body>
<h5>Loan History</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr>
		<th nowrap class="headrow" align="left" width="200">Loan Name</th>
		<th nowrap class="headrow" align="center">Date Requested</th>
		<th nowrap class="headrow" align="center">Date Approved</th>
		<th nowrap class="headrow" align="center">Loan Status</th>
		<th nowrap class="headrow" align="center">Loan Due Date</th>
    </tr>
<% 
while (!rsLoan.EOF) { 
%>
    <tr> 
		<td valign="top"><a href="javascript: openWindow('../LN/m008FS3.asp?intLoan_Req_id=<%=(rsLoan.Fields.Item("intLoan_Req_id").Value)%>','');"><%=(rsLoan.Fields.Item("chvLoan_name").Value)%></a></td>	
		<td valign="top" align="center"><%=FilterDate(rsLoan.Fields.Item("dtsRequest_date").Value)%></td>			
		<td valign="top" align="center"><%=FilterDate(rsLoan.Fields.Item("dtsApprvd_Date").Value)%></td>			
		<td valign="top" align="center"><%=(rsLoan.Fields.Item("chvLoan_Status_id").Value)%></td>
		<td valign="top" align="center"><%=FilterDate(rsLoan.Fields.Item("dtsLoan_Due_Date").Value)%></td>
    </tr>
<%
	rsLoan.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsLoan.Close();
%>
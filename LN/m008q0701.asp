<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

if (rsLoan.Fields.Item("insEq_user_type").Value=="3") {
	var rsDisability = Server.CreateObject("ADODB.Recordset");
	rsDisability.ActiveConnection = MM_cnnASP02_STRING;
	rsDisability.Source = "{call dbo.cp_dsbty_doc("+rsLoan.Fields.Item("intEq_user_id").Value+")}";
	rsDisability.CursorType = 0;
	rsDisability.CursorLocation = 2;
	rsDisability.LockType = 3;
	rsDisability.Open();
	
	var rsEducation = Server.CreateObject("ADODB.Recordset");
	rsEducation.ActiveConnection = MM_cnnASP02_STRING;
	rsEducation.Source = "{call dbo.cp_edu_doc2("+rsLoan.Fields.Item("intEq_user_id").Value+", 0,2,0)}";
//	rsEducation.Source = "{call dbo.cp_edu_doc("+rsLoan.Fields.Item("intEq_user_id").Value+")}";
	rsEducation.CursorType = 0;
	rsEducation.CursorLocation = 2;
	rsEducation.LockType = 3;
	rsEducation.Open();

	var semester = "";
	var year = "";
	var courses = "";
	var type = "";
	var eligible = "";
	while (!rsEducation.EOF) {
		semester = rsEducation.Fields.Item("chvsemester").Value;
		year = rsEducation.Fields.Item("insYear").Value;
		courses = rsEducation.Fields.Item("insNum_of_Course").Value;
		type = rsEducation.Fields.Item("chvCrse_type").Value;
		eligible = ((rsEducation.Fields.Item("bitIsElgb4_ASP").Value=="1")?"Yes":"No");
		rsEducation.MoveNext();
	}
	rsEducation.Requery();

	var rsExternalAgency = Server.CreateObject("ADODB.Recordset");
	rsExternalAgency.ActiveConnection = MM_cnnASP02_STRING;
	rsExternalAgency.Source = "{call dbo.cp_Ext_FS("+rsLoan.Fields.Item("intEq_user_id").Value+")}";
	rsExternalAgency.CursorType = 0;
	rsExternalAgency.CursorLocation = 2;
	rsExternalAgency.LockType = 3;
	rsExternalAgency.Open();

	var rsLoanOwnForm = Server.CreateObject("ADODB.Recordset");
	rsLoanOwnForm.ActiveConnection = MM_cnnASP02_STRING;
	rsLoanOwnForm.Source = "{call dbo.cp_Idv_LoanOwn_Form("+rsLoan.Fields.Item("intEq_user_id").Value+")}";
	rsLoanOwnForm.CursorType = 0;
	rsLoanOwnForm.CursorLocation = 2;
	rsLoanOwnForm.LockType = 3;
	rsLoanOwnForm.Open();
}
%>
<html>
<head>
	<title>Documentation & Eligibility</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>
</head>
<body>
<h5>Documentation & Eligibility</h5>
<%
if (rsLoan.Fields.Item("insEq_user_type").Value=="3") {
%>
<b>Disability Documentation:</b>
<table cellpadding="2" cellspacing="0" style="border: 1px solid" width="80%">
	<tr>
		<td style="border: 1px solid">Type</td>
		<td style="border: 1px solid">Date</td>
		<td style="border: 1px solid" align="center">Eligible</td>
	</tr>
<%
while (!rsDisability.EOF){	
%>
    <tr> 
		<td style="border: 1px solid"><%=(rsDisability.Fields.Item("chvType").Value)%></td>
		<td style="border: 1px solid"><%=FilterDate(rsDisability.Fields.Item("dtsDocDate").Value)%></td>
		<td style="border: 1px solid" align="center"><%=rsDisability.Fields.Item("bitEligible").Value%></td>
    </tr>
<%
	rsDisability.MoveNext();
}
%>
</table>
<br><br>
<b>Education Documentation:</b>
<table cellpadding="2" cellspacing="0" style="border: 1px solid" width="80%">
	<tr>
		<td style="border: 1px solid" align="left">Semester</td>
		<td style="border: 1px solid" align="left">Year</td>
		<td style="border: 1px solid" align="center">Number Of Courses</td>
		<td style="border: 1px solid" align="left">Course Type</td>
		<td style="border: 1px solid" align="center">Eligible</td>		
	</tr>
<%
if (!rsEducation.EOF) {
%>
	<tr>
		<td style="border: 1px solid" align="left"><%=semester%>&nbsp;</td>
		<td style="border: 1px solid" align="left"><%=year%>&nbsp;</td>
		<td style="border: 1px solid" align="center"><%=courses%>&nbsp;</td>
		<td style="border: 1px solid" align="left"><%=type%>&nbsp;</td>
		<td style="border: 1px solid" align="center"><%=eligible%>&nbsp;</td>		
	</tr>
<%
}
%>
</table>
<br><br>
<b>External Agency:</b>
<table cellpadding="2" cellspacing="0" style="border: 1px solid" width="80%">
	<tr>
		<td style="border: 1px solid" align="left">Type</td>
		<td style="border: 1px solid" align="left">Claim Status</td>
		<td style="border: 1px solid" align="center">Eligible for CSG</td>
		<td style="border: 1px solid" align="center">Eligible for EPPD</td>		
	</tr>
<%
while (!rsExternalAgency.EOF){	
%>
	<tr>
		<td style="border: 1px solid" align="left"><%=(rsExternalAgency.Fields.Item("chvExtFS_type").Value)%></td>
		<td style="border: 1px solid" align="left"><%=(rsExternalAgency.Fields.Item("chvClaim_Status").Value)%></td>
		<td style="border: 1px solid" align="center"><%=(rsExternalAgency.Fields.Item("chvElgb4_ASP").Value)%></td>
		<td style="border: 1px solid" align="center"><%=(rsExternalAgency.Fields.Item("chrElgb_VR").Value)%></td>
	</tr>
<%
	rsExternalAgency.MoveNext();
}
%>
</table>
<br><br>
<% 
if (!rsLoanOwnForm.EOF) {
%>
Is First Nation: <%=((rsLoanOwnForm.Fields.Item("bitIs_FirstNations").Value=="1")?"Yes":"No")%>
<br>
<br>
<%
}
%>
<b>Loan Own Form:</b>
<table cellpadding="2" cellspacing="0" style="border: 1px solid" width="40%">
	<tr>
		<td style="border: 1px solid" align="center">Date Received</td>
	</tr>
<%
while (!rsLoanOwnForm.EOF){	
%>
	<tr>
		<td style="border: 1px solid" align="center"><%=FilterDate(rsLoanOwnForm.Fields.Item("dtsLOform_rx_date").Value)%></td>
	</tr>
<%
	rsLoanOwnForm.MoveNext();
}
%>
</table>
<%
} else {
%>
<i>Information not available for this loan.</i>
<%
}
%>
</body>
</html>
<%
if (rsLoan.Fields.Item("insEq_user_type").Value=="3") {
	rsLoanOwnForm.Close();
	rsDisability.Close();
	rsExternalAgency.Close();
	rsEducation.Close();
}
rsLoan.Close();
%>
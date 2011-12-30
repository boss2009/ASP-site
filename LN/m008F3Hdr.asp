<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var rsLoanHeader = Server.CreateObject("ADODB.Recordset");
rsLoanHeader.ActiveConnection = MM_cnnASP02_STRING;
rsLoanHeader.Source = "{call dbo.cp_FrmHdr_8("+Request.QueryString("intLoan_Req_id")+")}";
rsLoanHeader.CursorType = 0;
rsLoanHeader.CursorLocation = 2;
rsLoanHeader.LockType = 3;
rsLoanHeader.Open();
%>
<html>
<head>
	<title>Loan Request Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 570px"> 
<%
if (rsLoanHeader.EOF) {
%>
<i>Information not available for this loan.</i> <br>
<br>
<br>
<br>
<%
} else {
	switch (rsLoan.Fields.Item("insEq_user_type").Value) {
		case 2:
%>
<i>Information not available for this SETBC loan.</i> <br>
<br>
<br>
<br>
<%
	break;
	case 3:
%>
  <table cellspacing="1" cellpadding="1">
    <tr> 
      <td><b>User Name:</b></td>
      <td width="200"><%=(rsLoanHeader.Fields.Item("chvEq_user_Name").Value)%></td>
      <td><b>Disability:</b></td>
      <td><%=(rsLoanHeader.Fields.Item("chvDisability").Value)%></td>
    </tr>
    <tr> 
      <td><b>User Type:</b></td>
      <td><%=(rsLoanHeader.Fields.Item("chvEq_user_type").Value)%></td>
      <td><b>Case Manager:</b></td>
      <td><%=(rsLoanHeader.Fields.Item("chvCaseManager").Value)%></td>
    </tr>
    <tr> 
      <td valign="top"><b>Address:</b></td>
      <td valign="top"> <%=(rsLoanHeader.Fields.Item("chvAddress").Value)%><br>
        <%=(rsLoanHeader.Fields.Item("chvCity").Value)%>&nbsp;<%=(rsLoanHeader.Fields.Item("chrprvst_abbv").Value)%>&nbsp;<%=(rsLoanHeader.Fields.Item("chvPostal_zip").Value)%>&nbsp;</td>
      <td valign="top"><b>Loan Type:</b></td>
      <td valign="top"><%=(rsLoanHeader.Fields.Item("chvLoan_Type").Value)%></td>
    </tr>
    <tr> 
      <td><b>Phone Number:</b></td>
      <td><%=FormatPhoneNumber(rsLoanHeader.Fields.Item("chvPhone_Type_1").Value,rsLoanHeader.Fields.Item("chvPhone1_Arcd").Value,rsLoanHeader.Fields.Item("chvPhone1_Num").Value,rsLoanHeader.Fields.Item("chvPhone1_Ext").Value,rsLoanHeader.Fields.Item("chvPhone_Type_2").Value,rsLoanHeader.Fields.Item("chvPhone2_Arcd").Value,rsLoanHeader.Fields.Item("chvPhone2_Num").Value,rsLoanHeader.Fields.Item("chvPhone2_Ext").Value,rsLoanHeader.Fields.Item("chvPhone_Type_3").Value,rsLoanHeader.Fields.Item("chvPhone3_Arcd").Value,rsLoanHeader.Fields.Item("chvPhone3_Num").Value,rsLoanHeader.Fields.Item("chvPhone3_Ext").Value)%></td>
      <td></td>
      <td></td>
    </tr>
  </table>
<%
	break;
	case 4:
%>
  <table cellspacing="1" cellpadding="1">
    <tr> 
      <td><b>Institution Name:</b></td>
      <td width="200"><%=(rsLoanHeader.Fields.Item("chvSchool_Name").Value)%></td>
      <td><b>Referral Date:</b></td>
      <td><%=FilterDate(rsLoanHeader.Fields.Item("dtsRefral_date").Value)%></td>
    </tr>
    <tr> 
      <td><b>Referring Agent:</b></td>
      <td><%=(rsLoanHeader.Fields.Item("chvReferring_Agent").Value)%></td>
      <td><b>Case Manager:</b></td>
      <td><%=(rsLoanHeader.Fields.Item("chvCase_mngr").Value)%></td>
    </tr>
    <tr> 
      <td valign="top"><b>Loan Type:</b></td>
      <td valign="top" colspan="3"><%=(rsLoanHeader.Fields.Item("chvLoan_Type").Value)%></td>
    </tr>
  </table>
<%
	break;
	case 1:
%>
  <table cellspacing="1" cellpadding="1">
    <tr> 
      <td><b>User Name:</b></td>
      <td width="160"><%=(rsLoanHeader.Fields.Item("chvEq_user_Name").Value)%></td>
      <td><b>Disability:</b></td>
      <td><%=(rsLoanHeader.Fields.Item("chvDisability").Value)%></td>
    </tr>
    <tr> 
      <td><b>User Type:</b></td>
      <td><%=(rsLoanHeader.Fields.Item("chvEq_user_type").Value)%></td>
      <td><b>Case Manager:</b></td>
      <td><%=(rsLoanHeader.Fields.Item("chvCaseManager").Value)%></td>
    </tr>
    <tr> 
      <td valign="top"><b>Address:</b></td>
      <td valign="top"> <%=(rsLoanHeader.Fields.Item("chvAddress").Value)%><br>
        <%=(rsLoanHeader.Fields.Item("chvCity").Value)%>&nbsp;<%=(rsLoanHeader.Fields.Item("chrprvst_abbv").Value)%>&nbsp;<%=(rsLoanHeader.Fields.Item("chvPostal_zip").Value)%></td>
      <td valign="top"><b>Loan Type:</b></td>
      <td valign="top"><%=(rsLoanHeader.Fields.Item("chvLoan_Type").Value)%></td>
    </tr>
    <tr> 
      <td nowrap valign="top"><b>Phone Number:</b></td>
      <td width="140"><%=FormatPhoneNumber(rsLoanHeader.Fields.Item("chvPhone_Type_1").Value,rsLoanHeader.Fields.Item("chvPhone1_Arcd").Value,rsLoanHeader.Fields.Item("chvPhone1_Num").Value,rsLoanHeader.Fields.Item("chvPhone1_Ext").Value,rsLoanHeader.Fields.Item("chvPhone_Type_2").Value,rsLoanHeader.Fields.Item("chvPhone2_Arcd").Value,rsLoanHeader.Fields.Item("chvPhone2_Num").Value,rsLoanHeader.Fields.Item("chvPhone2_Ext").Value,rsLoanHeader.Fields.Item("chvPhone_Type_3").Value,rsLoanHeader.Fields.Item("chvPhone3_Arcd").Value,rsLoanHeader.Fields.Item("chvPhone3_Num").Value,rsLoanHeader.Fields.Item("chvPhone3_Ext").Value)%></td>
      <td nowrap valign="top"></td>
      <td></td>
    </tr>
  </table>
<%
	break;
	}
}
%>
</div>
</body>
</html>
<%
rsLoan.Close();
rsLoanHeader.Close();
%>
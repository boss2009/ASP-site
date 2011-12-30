<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ClientLetterHeader.inc"-->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Loan - Pending Buyout</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><img src="http://<%=Request.ServerVariables("server_name")%>:8080/i/letterhead.gif" width="450" height="80"></p>
<br>
<%=Creation_Date%><br>
<br>
<%=Recipient_Whole_Name%><br>
<%=Recipient_SIN%><br>
<%=Recipient_School_Address%><br>
<br>
Dear <%=Recipient_Title%>&nbsp;<%=Recipient_Last_Name%>,<br>
<br>
Re: Buyout Cost of Loaned Equipment<br>
<%
if (Is_Employment_Buyout) {
%>
<p>
As per EPPD policy, Sirius Innovations Inc. supports clients with loaned
equipment in their work environment for approximately <%=Employment_Loan_Duration%> to assess the
suitability of the adaptive equipment in obtaining and/or maintaining employment.
</p>
<p>
At this time, your present loan period has expired as of <%=Loan_Expiry_Date%> and the
loaned equipment will need to be purchased or returned as per Ministry policy.
After reviewing the options, you have agreed to the following purchase plan:
</p>
<%
} else {
%>
<p>
As per our conversation about your current loaned equipment, you have
decided that you would like to purchase the loaned equipment and have agreed
to the following purchase plan:
</p>
<%
}
%>
<p>
Equipment List and Original Cost
<%=Equipment_Still_On_Loan_List_With_Cost%>
</p>
<p>
Sub-total: <%=Equipment_Still_On_Loan_List_Total_Cost%><br>
Less Discount: <%=Discount_Amount%><br>
Buyout Cost: <%=Buyout_Cost%><br>
</p>
<p>
Number of Installments: <%=Number_Of_Installments%><br>
<%=Installment_Due_Dates%><br>
Payment in Full Date: <%=Payment_In_Full_Date%><br>
</p>
<p>
Please send the cheque (payable to Sirius Innovations Inc.) to the address
listed on our letter.  Once we receive full payment, we will forward you a copy of
the invoice, stamped paid, as your proof of purchase.  Your file with Si2 will then
be closed.
</p>
<p>
Thank you for your attention to this matter.  If you have any questions or
concerns, do not hesitate to call.
</p>
Sincerely,<br>
<br>
<br>
<br>
D T Chan,<br> 
CEO
<br>
<br>
<%=CC_Whole_Name%>
</body>
</html>

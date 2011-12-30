<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/LoanLetterHeader.inc"-->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Loan - Accept</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><img src="http://<%=Request.ServerVariables("server_name")%>:8080/i/letterhead.gif" width="450" height="80"></p>
<br>
<%=Creation_Date%><br>
<br>
<%=Recipient_Title%>&nbsp;<%=Recipient_Whole_Name%><br>
<%=Recipient_Job_Position%><br>
<%=Recipient_Work_Address%><br>
<br>
Dear <%=Recipient_Title%>&nbsp;<%=Recipient_Last_Name%>,<br>
<br>
Re: <%=Subject_Whole_Name%>&nbsp;<%=Subject_SIN%><br>
<p>
As Manager of Sirius Innovations Inc. I am writing you with regards to
<%=Subject_Whole_Name%>'s referral application received by this office.  We are
pleased to inform you that <%=Subject_First_Name%> has been accepted for a loan of
the following equipment:</p>
<%=Loaned_Equipment_List%>
<%=Loan_Conditions%>
<%=Document_Conditions%>
<%=Training_Requested%>
<p>
If you have any further questions or concerns regarding this decision, please feel
free to contact me by calling (604) 959-8188.
</p>
<p>
Thank you for this referral and we hope that Si2 can be
of assistance to you in the future.
</p>
Yours truly,<br>
<br>
<br>
<br>
D T Chan	,<br> 
CEO
<br>
<br>
<%=CC_Whole_Name%>
</body>
</html>

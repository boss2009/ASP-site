<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ClientLetterHeader.inc"-->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Loan - Default</title>
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
&nbsp;&nbsp;&nbsp;&nbsp;DEFAULT STATUS
<p>
<%=Subject_Whole_Name%> received an equipment loan from Sirius Innovations Inc. for the purpose of <%=Purpose_Of_Loan%>.  The following list identifies 
the Si2 equipment on loan to <%=Subject_First_Name%>:
</p>
<%=Equipment_Still_On_Loan_List%>
<p>
At this point, <%=Subject_First_Name%> is no longer eligible for Si2 services due to
the following reason(s):
</p>
<%=Reason_For_Ineligibility%>
<p>
Si2 has been unable to retrieve the equipment despite numerous attempts to retrieve
the equipment; therefore, Si2 will place <%=Subject_Whole_Name%> in default of the
established equipment loan agreement plan.  Si2 will record this status in our 
database and close <%=Subject_First_Name%>'s file.
</p>
<p>
In addition, Si2 is requesting that this letter be placed in your organization's/
institution's file as a reminder that <%=Subject_First_Name%> is not in good standing
with Si2.
</p>
<p>
In order to rescind the default status, we require arrangements to be made for either
the return of equipment, or payment in the amount of <%=Purchase_Cost_Of_Equipment%>
to purchase the outstanding loaned equipment.
</p>
<p>
Please call me if you wish to discuss this matter further.
</p>
Sincerely,<br>
<br>
<br>
<br>
D T Chan,<br>
CEO<br>
<br>
<%=CC_Whole_Name%>
</body>
</html>

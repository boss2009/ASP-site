<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/BuyoutLetterHeader.inc"-->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>CSG - Accept</title>
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
<p>
As Manager of Sirius Innovations Inc., I am writing you 
with regards to your TSSP application received by this office.  We are pleased 
to inform you that you are eligible for the Canada Study Grant for Students with 
Permanent Disabilities by demonstrating financial need through <%=Grant_Qualification_Source%>.  
This funding is to enable you to purchase equipment during your study period from
<%=Study_Period_From%> to <%=Study_Period_To%>.</p>
<p>
The details of the agreed upon technology plan are as follows:<br>
- Use eligible funds to purchase the following equipment:<br>
<%=Sold_Equipment_List%>
</p>
<%=Return_For_Donation%>
<%=Return_For_Loan%>
<%=Ship_For_Configuration_Requested%>

<%=Document_Conditions%>
<p>
Please be advised that the equipment purchased through CSG is subject to federal
income tax regulations.
</p>
<%=Conditions%>
<%=Not_Donation_Configuration_Loan_Return%>
<%=Donation%>
<%=Loan_Return%>
<%=Configuration_Requested%>
<%=Shipping_Origin%>
<%=Training_Requested%>
<p>
If you have any further questions or concerns regarding this decision, please feel
free to contact me by calling (604) 959-8188.
</p>
<p>
Thank you for applying to Si2 and we hope that Sirius Innovations Inc. can be
of assistance to you in the future.
</p>
Yours truly,<br>
<br>
<br>
D T Chan,<br> 
CEO
<br>
<br>
<%=CC_Whole_Name%>
</body>
</html>

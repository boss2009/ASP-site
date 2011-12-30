<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ClientLetterHeader.inc"-->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Loan - Rescind Default</title>
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
Further to our letter dated <%=Last_Loan_Default_Letter%> regarding <%=Subject_Whole_Name%>'s
default status with Sirius Innovations Inc., we have now received <%=Reason_For_Canceling_Default%>.
</p>
<p>
Accordingly, Si2 will rescind <%=Subject_First_Name%>'s default status with our program
and make a note in our database of <%=Subject_First_Name%>'s good standing.
</p>
<p>
Please call me if you wish to discuss this further.
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

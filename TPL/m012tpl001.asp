<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/PILATLetterHeader.inc"-->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>PILAT - Decline</title>
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
Re: <%=Pilat_Decline_Referral_Type%><br>
<p>
I am writing you with regards to your PILAT referral application received by Sirius Innovations Inc.</p>
<p>
At this time, the application is being put on hold for the following reason(s):
</p>
<%=Decline_Reasons%>
<p>
Si2 will re-activate the file when the above issues have been resolved.
</p>
<p>
If you have any further questions or concerns regarding this decision, please feel
free to contact me by calling (604) 959-8188. </p>
<p>
Thank you for the referral and we hope that Si2 can be of assistance to you in the future.
</p>
Yours truly,<br>
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

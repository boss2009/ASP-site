<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ClientLetterHeader.inc"-->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>CSG - MIR</title>
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
<p>
As Manager of Sirius Innovations Inc., I am writing you with 
regards to your TSSP application received by this office.  To be eligible for the 
Canada Study Grant (CSG) Program for Students with Permanent Disabilities, you must 
have a permanent disability, be taking a post secondary level course at a designated
institution either full or part-time, and demonstrate financial need through either a
British Columbia Student Assistance Program (BCSAP) or High-Need Part-Time (HNPT) CSG
Program.
</p>
<%=CSG_Issues%>
<%=CSG_Missing_Documentation%>
<p>
If you have any further questions or concerns regarding this decision, please feel free
to contact me by calling (604) 959-8188.
</p>
<p>
Thank you for applying and we hope that the Sirius Innovations Inc. can be of
assistance to you in the future.
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

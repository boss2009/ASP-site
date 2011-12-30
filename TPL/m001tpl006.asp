<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ClientLetterHeader.inc"-->
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Loan - Annual Education Follow-Up</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><img src="http://<%=Request.ServerVariables("server_name")%>:8080/i/letterhead.gif" width="450" height="80"></p>
<br>
<%=Creation_Date%><br>
<br>
<%=Recipient_Whole_Name%><br>
<%=Recipient_Job_Position%><br>
<%=Recipient_Work_Address%><br>
<br>
Dear <%=Recipient_Title%>&nbsp;<%=Recipient_Last_Name%>,<br>
<br>
Re: <%=Subject_Whole_Name%><br>
Annual Educational Follow-Up&nbsp;<%=CurrentYear()%><br>
<p>
We are now at the point of initiating our annual follow-up process for students
who have equipment on loan from Sirius Innovations Inc..  <%=Subject_Whole_Name%>
currently has the following equpment on loan:
</p>
<%=Equipment_Still_On_Loan_List%>
<p>
As per our contract with the Ministry of Human Resources, Si2 is required to
initiate this process in order to:
</p>
<ul>
<li>ensure that the loaned equipment is being used by students to prepare for,
obtain, and maintain employment;
<li>arrange for the return of equipment to Si2 in a timely fashion
<li>process the high number of equipment requests with limited funding.
</ul>
<p>
In order for <%=Subject_First_Name%> to maintain the loan of the above equipment
over the summer and/or for the upcoming fall semester, <%=Subject_First_Name%> must
meet the following criteria as per Ministry guidelines:
</p>
<%=Conditions_To_Maintain_Loan%>
<p>
Therefore, please answer the following questions re: <%=Subject_First_Name%>:
</p>
<ol>
<li>What is the expected completion date of <%=Subject_First_Name%>'s current
educational program?
<li>Has <%=Subject_First_Name%> successfully completed his courses from January to April semester?
<li>Will <%=Subject_First_Name%> be enrolled in the minimum required courses for
the upcoming Fall semester and do you wish to continue the loan of the listed
equipment?
</ol>
<p>
If you are requesting special consideration with unsuccessful
course completion, please outline the educational plan to ensure that successful
course completion will be achieved in the coming semester.
</p>
<p>
Please email, fax or mail your reply to Si2 by <%=Reply_By_Date%>.
</p>
<p>
We appreciate your cooperation with this follow-up process.
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

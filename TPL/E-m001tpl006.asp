<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
If Request.Form("MM_send") <> "" Then
	on error resume next 'This code will only work on a Win2k server.
	Dim iMsg, iConf, Flds
	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	
	With Flds
	  ' assume constants are defined within script file
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.setbc.org"
	  .Update
	End With
	
	With iMsg
	  Set .Configuration = iConf
		  .To       = Request.Form("Recipient")
		  .Cc       = Request.Form("CC")
		  .From     = Request.Form("Sender")
		  .Subject  = Request.Form("Subject")
		  .TextBody = Request.Form("Message")
		  .Send
	End With
	Response.Redirect("../AC/InsertSuccessful.html")	
End If
</script>
<!--#include file="../inc/ClientLetterHeader.inc"-->
<html>
<head>
	<title>Loan - Annual Education Follow-Up</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Send(){
		if (!CheckEmail(document.frm0106.Recipient.value)) {
			alert("Invalid recipient email address.");
			document.frm0106.Recipient.focus();
			return ;
		}
		if (Trim(document.frm0106.Recipient.value) == "") {
			alert("Enter recipient email address.");
			document.frm0106.Recipient.focus();
			return ;
		}
		if (!CheckEmail(document.frm0106.Sender.value)) {
			alert("Invalid sender email address.");
			document.frm0106.Sender.focus();
			return ;
		}
		if (Trim(document.frm0106.Sender.value) == "") {
			alert("Enter sender email address.");
			document.frm0106.Sender.focus();
			return ;
		}
		document.frm0106.submit();
	}
	
	function Init(){
		document.frm0106.Recipient.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0106" method="POST">
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Recipient:</td>
		<td><input type="text" name="Recipient" tabindex="1" value="<%=Recipient_Email%>" accesskey="F"></td>
	</tr>
	<tr>
		<td>CC:</td>
		<td><input type="text" name="CC" tabindex="2" value="<%=CC_Email%>"></td>
	</tr>
	<tr>
		<td>Sender:</td>
		<td><input type="text" name="Sender" tabindex="3" value="<%=Sender_Email%>"></td>
	</tr>
	<tr>
		<td>Subject:</td>
		<td><input type="text" name="Subject" size="75" value="Re: <%=Subject_Whole_Name%> Annual Educational Follow-Up <%=CurrentYear()%>" tabindex="4">
	</tr>
<tr>		
<td valign="top">Message:</td>
<td valign="top"><textarea name="Message" cols="75" rows="15" tabindex="5" accesskey="L">
Dear <%=Recipient_Title%>&nbsp;<%=Recipient_Last_Name%>,

We are now at the point of initiating our annual follow-up process for students who have equipment on loan from Sirius Innovations Inc.  <%=Subject_Whole_Name%> currently has the following equpment on loan:

<%=ReplaceTags(Equipment_Still_On_Loan_List)%>

As per our contract with the Ministry of Human Resources, Si2 is required to initiate this process in order to:

-ensure that the loaned equipment is being used by students to prepare for, obtain, and maintain employment;
-arrange for the return of equipment to Si2 in a timely fashion
-process the high number of equipment requests with limited funding.

In order for <%=Subject_First_Name%> to maintain the loan of the above equipment over the summer and/or for the upcoming fall semester, <%=Subject_First_Name%> must meet the following criteria as per Ministry guidelines:

<%=ReplaceTags(Conditions_To_Maintain_Loan)%>

Therefore, please answer the following questions re: <%=Subject_First_Name%>:

1) What is the expected completion date of <%=Subject_First_Name%>'s current educational program?
2) Has <%=Subject_First_Name%> successfully completed his courses from January to April semester?
3) Will <%=Subject_First_Name%> be enrolled in the minimum required courses for the upcoming Fall semester and do you wish to continue the loan of the listed equipment?

If you are requesting special consideration with unsuccessful course completion, please outline the educational plan to ensure that successful course completion will be achieved in the coming semester.

Please email, fax or mail your reply to Si2 by <%=Reply_By_Date%>.

We appreciate your cooperation with this follow-up process.

Yours truly,


D T Chan,
CEO

<%=ReplaceTags(CC_Whole_Name)%>
</textarea></td>
</tr>
</table>
<input type="button" value="Send" class="btnstyle" onClick="Send();" tabindex="6">
<input type="button" value="Cancel" class="btnstyle" onClick="window.close();" tabindex="7">
<input type="hidden" name="MM_send" value="1">
</form>
</body>
</html>
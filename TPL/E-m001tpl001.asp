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
	<title>Loan - Default</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Send(){
		if (!CheckEmail(document.frm0101.Recipient.value)) {
			alert("Invalid recipient email address.");
			document.frm0101.Recipient.focus();
			return ;
		}
		if (Trim(document.frm0101.Recipient.value) == "") {
			alert("Enter recipient email.");
			document.frm0101.Recipient.focus();
			return ;
		}
		if (!CheckEmail(document.frm0101.Sender.value)) {
			alert("Invalid sender email address.");
			document.frm0101.Sender.focus();
			return ;
		}
		if (Trim(document.frm0101.Sender.value) == "") {
			alert("Enter sender email address.");
			document.frm0101.Sender.focus();
			return ;
		}
		document.frm0101.submit();
	}
	
	function Init(){
		document.frm0101.Recipient.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0101" method="POST">
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
		<td><input type="text" name="Subject" size="75" value="Re: <%=Subject_Whole_Name%> <%=Subject_SIN%> - DEFAULT STATUS" tabindex="4">
	</tr>
<tr>		
<td valign="top">Message:</td>
<td valign="top">
<textarea name="Message" cols="75" rows="15" tabindex="5" accesskey="L">
Dear <%=Recipient_Title%>&nbsp;<%=Recipient_Last_Name%>,

<%=Subject_Whole_Name%> received an equipment loan from Sirius Innovations Inc. for the purpose of <%=Purpose_Of_Loan%>.  The following list identifies the Si2 equipment on loan to <%=Subject_First_Name%>:
<%=ReplaceTags(Equipment_Still_On_Loan_List)%>

At this point, <%=Subject_First_Name%> is no longer eligible for Si2 services due to the following reason(s):
<%=ReplaceTags(Reason_For_Ineligibility)%>

Si2 has been unable to retrieve the equipment despite numerous attempts to retrieve the equipment; therefore, Si2 will place <%=Subject_First_Name%> in default of the established equipment loan agreement plan.  Si2 will record this status in our database and close <%=Subject_First_Name%>'s file.

In addition, Si2 is requesting that this letter be placed in your organization's/institution's file as a reminder that <%=Subject_First_Name%> is not in good standing with Si2.

In order to rescind the default status, we require arrangements to be made for either the return of equipment, or payment in the amount of <%=Purchase_Cost_Of_Equipment%> to purchase the outstanding loaned equipment.

Please call me if you wish to discuss this matter further.

Sincerely,


D T Chan,
CEO

<%=ReplaceTags(CC_Whole_Name)%>
</textarea>
</td>
</tr>
</table>
<input type="button" value="Send" class="btnstyle" onClick="Send();" tabindex="6">
<input type="button" value="Cancel" class="btnstyle" onClick="window.close();" tabindex="7">
<input type="hidden" name="MM_send" value="1">
</form>
</body>
</html>
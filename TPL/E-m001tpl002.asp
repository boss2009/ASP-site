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
	<title>Loan - Rescind Default</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Send(){
		if (!CheckEmail(document.frm0102.Recipient.value)) {
			alert("Invalid Recipient Email.");
			document.frm0102.Recipient.focus();
			return ;
		}
		if (document.frm0102.Recipient.value == "") {
			alert("Missing Recipient Email.");
			document.frm0102.Recipient.focus();
			return ;
		}
		if (!CheckEmail(document.frm0102.Sender.value)) {
			alert("Invalid Sender Email.");
			document.frm0102.Sender.focus();
			return ;
		}
		if (document.frm0102.Sender.value == "") {
			alert("Missing Sender email.");
			document.frm0102.Sender.focus();
			return ;
		}
		document.frm0102.submit();
	}
	
	function Init(){
		document.frm0102.Recipient.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0102" method="POST">
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
		<td><input type="text" name="Subject" size="75" value="Re: <%=Subject_Whole_Name%>&nbsp;<%=Subject_SIN%> - DEFAULT STATUS" tabindex="4">
	</tr>
<tr>		
<td valign="top">Message:</td>
<td valign="top"><textarea name="Message" cols="75" rows="15" tabindex="5" accesskey="L">
Dear <%=Recipient_Title%>&nbsp;<%=Recipient_Last_Name%>,

Further to our letter dated <%=Last_Loan_Default_Letter%> regarding <%=Subject_Whole_Name%>'s default status with Sirius Innovations Inc., we have now received <%=ReplaceTags(Reason_For_Canceling_Default)%>.

Accordingly, Si2 will rescind <%=Subject_First_Name%>'s default status with our program
and make a note in our database of <%=Subject_First_Name%>'s good standing.

Please call me if you wish to discuss this further.

Sincerely,


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
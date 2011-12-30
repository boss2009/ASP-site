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
<!--#include file="../inc/PILATLetterHeader.inc"-->
<html>
<head>
	<title>Student- Accept</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Send(){
		if (!CheckEmail(document.frm1202.Recipient.value)) {
			alert("Invalid recipient email address.");
			document.frm1202.Recipient.focus();
			return ;
		}
		if (Trim(document.frm1202.Recipient.value) == "") {
			alert("Enter recipient email address.");
			document.frm1202.Recipient.focus();
			return ;
		}
		if (!CheckEmail(document.frm1202.Sender.value)) {
			alert("Invalid sender email address.");
			document.frm1202.Sender.focus();
			return ;
		}
		if (Trim(document.frm1202.Sender.value) == "") {
			alert("Enter sender email address.");
			document.frm1202.Sender.focus();
			return ;
		}
		document.frm1202.submit();
	}
	
	function Init(){
		document.frm1202.Recipient.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm1202" method="POST">
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
		<td><input type="text" name="Subject" size="75" value="Re: <%=Pilat_Accept_Referral_Type%>" tabindex="4">
	</tr>
<tr>		
<td valign="top">Message:</td>
<td valign="top"><textarea name="Message" cols="75" rows="15" tabindex="5" accesskey="L">
Dear <%=Recipient_Title%>&nbsp;<%=Recipient_Last_Name%>,

I am writing you with regards to your PILAT referral application received by Sirius Innovations Inc.

<%=ReplaceTags(Introduction)%>
<%=ReplaceTags(Equipment_List)%>
<%=ReplaceTags(Equipment_Conditions)%>
<%=ReplaceTags(Document_Conditions)%>
<%=ReplaceTags(Training_Requested)%>

If you have any further questions or concerns regarding this decision, please feel free to contact me by calling (604) 959-8188.

Thank you for the referral and we hope that Si2 can be of assistance to you in the future.

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
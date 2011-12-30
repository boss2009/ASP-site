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
	<title>Loan - Pending Buyout</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Send(){
		if (!CheckEmail(document.frm0103.Recipient.value)) {
			alert("Invalid recipient email address.");
			document.frm0103.Recipient.focus();
			return ;
		}
		if (Trim(document.frm0103.Recipient.value) == "") {
			alert("Enter recipient email address.");
			document.frm0103.Recipient.focus();
			return ;
		}
		if (!CheckEmail(document.frm0103.Sender.value)) {
			alert("Invalid sender email address.");
			document.frm0103.Sender.focus();
			return ;
		}
		if (Trim(document.frm0103.Sender.value) == "") {
			alert("Enter sender email address.");
			document.frm0103.Sender.focus();
			return ;
		}
		document.frm0103.submit();
	}
	
	function Init(){
		document.frm0103.Recipient.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0103" method="POST">
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
		<td><input type="text" name="Subject" size="75" value="Re: Buyout Cost of Loaned Equipment" tabindex="4">
	</tr>
<tr>		
<td valign="top">Message:</td>
<td valign="top"><textarea name="Message" cols="75" rows="15" tabindex="5" accesskey="L">
Dear <%=Recipient_Title%>&nbsp;<%=Recipient_Last_Name%>,
<%
if (Is_Employment_Buyout) {
%>
As per EPPD policy, Sirius Innovations Inc. supports clients with loaned equipment in their work environment for approximately <%=Employment_Loan_Duration%> to assess the suitability of the adaptive equipment in obtaining and/or maintaining employment.

At this time, your present loan period has expired as of <%=Loan_Expiry_Date%> and the loaned equipment will need to be purchased or returned as per Ministry policy.  After reviewing the options, you have agreed to the following purchase plan:
<%
} else {
%>
As per our conversation about your current loaned equipment, you have decided that you would like to purchase the loaned equipment and have agreed to the following purchase plan:
<%
}
%>
Equipment List and Original Cost
<%=ReplaceTags(Equipment_Still_On_Loan_List_With_Cost)%>

Sub-total: <%=Equipment_Still_On_Loan_List_Total_Cost%>
Less Discount: <%=Discount_Amount%>
Buyout Cost: <%=Buyout_Cost%>

Number of Installments: <%=Number_Of_Installments%>
<%=ReplaceTags(Installment_Due_Dates)%>
Payment in Full Date: <%=Payment_In_Full_Date%>

Please send the cheque (payable to Sirius Innovations Inc.) to the address listed on our letter.  Once we receive full payment, we will forward you a copy of the invoice, stamped paid, as your proof of purchase.  Your file with Si2 will then be closed.

Thank you for your attention to this matter.  If you have any questions or concerns, do not hesitate to call.

Sincerely,


D T Chan,
CEO

<%=ReplaceTags(CC_Whole_Name)%>
</textarea></td>
</tr>
</table>
<input type="button" value="Send" class="btnstyle" onClick="Send();" tabindex="6">
<input type="button" value="Cancel" class="btnstyle" onClick="window.close();" tabindex="7">
</form>
</body>
</html>
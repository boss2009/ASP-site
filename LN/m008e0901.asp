<%@language="VBSCRIPT"%>
<!--#include file="../inc/VBLogin.inc"-->
<%
If Request.Form("MM_send") <> "" Then
	on error resume next 'This code will only work on a Win2k server.
	Dim iMsg, iConf, Flds
	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	
'	With Flds
'	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sirius-innovations.com"
'	  .Update
'	End With
' + Nov.06.2005
	Dim theSchema
	theSchema="http://schemas.microsoft.com/cdo/configuration/"
	set iConf=server.CreateObject("CDO.Configuration")
	iConf.fields.item(theSchema & "sendusing")=2
	iConf.fields.item(theSchema & "smtpserver")="mail.sirius-innovations.com"
	iConf.fields.Update

	With iMsg
	  Set .Configuration = iConf
		  .To       = Request.Form("Recipient")
		  .Cc       = Request.Form("CC")
		  .From     = Request.Form("Sender")
		  .Subject  = Request.Form("Subject")
		  .TextBody = Request.Form("Message")
		  .Send
	End With

    ' + Nov.06.2005	
	set iConf=Nothing
	Set iMsg =Nothing

End If

Dim rsLoan
Set rsLoan = Server.CreateObject("ADODB.Recordset")
rsLoan.ActiveConnection = MM_cnnASP02_STRING
rsLoan.Source = "{call dbo.cp_loan_request2(" & Request("intLoan_req_id") & ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}"
rsLoan.CursorType = 0
rsLoan.CursorLocation = 2
rsLoan.LockType = 3
rsLoan.Open()

Dim rsHeader
Set rsHeader = Server.CreateObject("ADODB.Recordset")
rsHeader.ActiveConnection = MM_cnnASP02_STRING
rsHeader.Source = "{call dbo.cp_FrmHdr_8(" & Request("intLoan_req_id") & ")}"
rsHeader.CursorType = 0
rsHeader.CursorLocation = 2
rsHeader.LockType = 3
rsHeader.Open()

dim intShip_dtl_id
dim insEq_user_type
intShip_dtl_id = 0
insEq_user_type = 0
If NOT rsLoan.EOF Then 
	If rsLoan.Fields.Item("intShip_dtl_id").Value > 0 Then
		intShip_dtl_id = rsLoan.Fields.Item("intShip_dtl_id").Value 
	End If
	insEq_user_type = rsLoan.Fields.Item("insEq_user_type").Value
End If

dim intEq_user_id
dim RefAgentEmail
dim RefAgentName
intEq_user_id = 0
If insEq_user_type = "3" Then
	intEq_user_id = rsLoan.Fields.Item("intEq_user_id").Value
	Dim rsContacts 
	Set rsContacts = Server.CreateObject("ADODB.Recordset")
	rsContacts.ActiveConnection = MM_cnnASP02_STRING
	rsContacts.Source = "{call dbo.cp_get_client_contact(" & intEq_user_id & ",0)}"
	rsContacts.CursorType = 0
	rsContacts.CursorLocation = 2
	rsContacts.LockType = 3
	rsContacts.Open
	While Not rsContacts.EOF
		If Trim(rsContacts.Fields.Item("chvRelationship").Value) = "Referring Agent" Then
			RefAgentName = rsContacts.Fields.Item("chvContact_Fst_name").Value
			RefAgentEmail = rsContacts.Fields.Item("chvHome_E_Mail").Value
			If RefAgentEmail = "" Then
				RefAgentEmail = Trim(rsContacts.Fields.Item("chvWork_E_Mail").Value)
			End If
		End If
		rsContacts.MoveNext()
	WEnd 
End If

Dim rsShippingAddress 
Set rsShippingAddress = Server.CreateObject("ADODB.Recordset")
rsShippingAddress.ActiveConnection = MM_cnnASP02_STRING
rsShippingAddress.Source = "{call dbo.cp_loan_ship_address(" & intShip_dtl_id &",0,'','','','','',0,'','',0,'',0,'','','',0,'','','',0,'','','','','',0,'Q',0)}"
rsShippingAddress.CursorType = 0
rsShippingAddress.CursorLocation = 2
rsShippingAddress.LockType = 3
rsShippingAddress.Open

Dim rsShippingMethod
Set rsShippingMethod = Server.CreateObject("ADODB.Recordset")
rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING
rsShippingMethod.Source = "{call dbo.cp_loan_ship_method("& intShip_dtl_id & ",0,'',0,0,0,'',0,'','',0,0,'Q',0)}"
rsShippingMethod.CursorType = 0
rsShippingMethod.CursorLocation = 2
rsShippingMethod.LockType = 3
rsShippingMethod.Open()

Dim rsSender
Set rsSender = Server.CreateObject("ADODB.Recordset")
rsSender.ActiveConnection = MM_cnnASP02_STRING
rsSender.Source = "{call dbo.cp_logmster(140,'','',0,1,'Q',0)}"
rsSender.CursorType = 0
rsSender.CursorLocation = 2
rsSender.LockType = 3
rsSender.Open()
%>
<html>
<head>
	<title>Email Referring Agent</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Send(){
		if (!CheckEmail(document.frm0901.Recipient.value)) {
			alert("Invalid Recipient Email.");
			document.frm0901.Recipient.focus();
			return ;
		}
		if (document.frm0901.Recipient.value == "") {
			alert("Missing Recipient Email.");
			document.frm0901.Recipient.focus();
			return ;
		}
		if (!CheckEmail(document.frm0901.Sender.value)) {
			alert("Invalid Sender Email.");
			document.frm0901.Sender.focus();
			return ;
		}
		if (document.frm0901.Sender.value == "") {
			alert("Missing Sender email.");
			document.frm0901.Sender.focus();
			return ;
		}
		document.frm0901.submit();
	}
	
	function Init(){
	<%
	If Not intShip_dtl_id = 0 and Not rsHeader.EOF Then
	%>	
		document.frm0901.Recipient.focus();
	<%
	End If
	%>
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0901" method="POST" action="m008e0901.asp?intLoan_req_id=<%=Request("intLoan_req_id")%>">
<%
If intShip_dtl_id = 0 Then
%>
<i>Not available without shipping information.</i>
<%
Elseif rsHeader.EOF Then
%>
<i>This loan belongs to Si2.</i>
<%
Else
%>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Recipient:</td>
		<td><input type="text" name="Recipient" tabindex="1" value="<%=RefAgentEmail%>" accesskey="F"></td>
	</tr>
	<tr>
		<td>CC:</td>
		<td><input type="text" name="CC" tabindex="2"></td>
	</tr>
	<tr>
		<td>Sender:</td>
		<td><input type="text" name="Sender" tabindex="3" value="<%=rsSender.Fields.Item("chvEmail").Value%>"></td>
	</tr>
	<tr>
		<td>Subject:</td>
		<td><input type="text" name="Subject" size="75" value="<%=(rsHeader.Fields.Item("chvEq_user_Name").Value)%>: Shipping Arrangement" tabindex="4">
	</tr>
<tr>		
<td valign="top">Message:</td>
<td valign="top"><textarea name="Message" cols="75" rows="15" tabindex="5" accesskey="L">
Hi <%=RefAgentName%>,
Please be advised we have arranged for the delivery of the Si2 loan equipment to <%=(rsHeader.Fields.Item("chvEq_user_Name").Value)%>.
<%
Select Case rsShippingMethod.Fields.Item("insShip_Method_id").Value
'Loomis
case "4":
%>		
We will be shipping the equipment from our warehouse on <%=rsShippingMethod.Fields.Item("dtsDlvy_date").Value%>.
<%
'Dynamex
case "9":
%>
We will be shipping the equipment from our warehouse on <%=rsShippingMethod.Fields.Item("dtsDlvy_date").Value%>.
<%
'Client Pickup
case "10":
%>
<%=(rsHeader.Fields.Item("chvEq_user_Name").Value)%> has opted to pick up the equipment on <%=rsShippingMethod.Fields.Item("dtsDlvy_date").Value%>.
<%
'Case Manager Delivery
case "1":
%>
<%=rsHeader.Fields.Item("chvCaseManager").Value%> is delivering the equipment on <%=rsShippingMethod.Fields.Item("dtsDlvy_date").Value%> so that the equipment can be set up and training on adaptive software conducted.
<%
End Select
%>
If you have any questions, please do not hesitate to call me at (604)959-8188.
		
Web Master

Sirius Innovation Inc
P.O.Box 43119 Richmond Ctr PO, Richmond, BC, V6V 2W4
O: (604)959-8188
</textarea></td>
</tr>
</table>
<!--
<input type="submit" value="Send" class="btnstyle" tabindex="6">
-->
<input type="submit" value="Send" class="btnstyle" tabindex="6">

<input type="hidden" name="MM_send" value="1">
<%
End If
%>
</form>
</body>
</html>
<%
Set rsSender = Nothing
Set rsLoan = Nothing
Set rsHeader = Nothing
Set rsShippingMethod = Nothing
Set rsShippingAddress = Nothing
%>
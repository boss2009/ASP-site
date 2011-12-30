<%@language="VBSCRIPT"%>
<!--#include file="../inc/VBLogin.inc"-->
<%
If Request.Form("MM_send") <> "" Then
	on error resume next 'This code will only work on a Win2k server.
	Dim iMsg, iConf, Flds
	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	With Flds
	  ' assume constants are defined within script file
	  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.sirius-innovations.com"
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'      .Item("http://schemas.microsoft.com/cdo/configuration/sendusing")  = CdoSendUsingPort

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
End If

Dim rsShippingMethod 
Set rsShippingMethod = Server.CreateObject("ADODB.Recordset")
rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING
rsShippingMethod.Source = "{call dbo.cp_eqpsrv_ship_method(" & Request.QueryString("intEquip_srv_id") & ",'',0,0,'','','',0,'',0,1,'Q',0)}"
rsShippingMethod.CursorType = 0
rsShippingMethod.CursorLocation = 2
rsShippingMethod.LockType = 3
rsShippingMethod.Open()

Dim rsEmailFields
Set rsEmailFields = Server.CreateObject("ADODB.Recordset")
rsEmailFields.ActiveConnection = MM_cnnASP02_STRING
rsEmailFields.Source = "{call dbo.cp_eqpsrv_email("& Request.QueryString("intEquip_srv_id") & ",0)}"
rsEmailFields.CursorType = 0
rsEmailFields.CursorLocation = 2
rsEmailFields.LockType = 3
rsEmailFields.Open()

Dim UserName 
UserName = ""
If Not rsEmailFields.EOF Then
	If rsEmailFields.Fields.Item("insUser_Type_id").Value = 3 Then
		UserName = Trim(rsEmailFields.Fields.Item("chvUsrFst_Name").Value) & " " & Trim(rsEmailFields.Fields.Item("chvUsrLst_Name").Value)
	Else 
		UserName = rsEmailFields.Fields.Item("chvSchool_Name").Value
	End If
End If

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
		if (!CheckEmail(document.frm0501.Recipient.value)) {
			alert("Invalid Recipient Email.");
			document.frm0501.Recipient.focus();
			return ;
		}
		if (Trim(document.frm0501.Recipient.value) == "") {
			alert("Missing Recipient Email.");
			document.frm0501.Recipient.focus();
			return ;
		}
		if (!CheckEmail(document.frm0501.Sender.value)) {
			alert("Invalid Sender Email.");
			document.frm0501.Sender.focus();
			return ;
		}
		if (Trim(document.frm0501.Sender.value) == "") {
			alert("Missing Sender Email.");
			document.frm0501.Sender.focus();
			return ;
		}
		document.frm0501.submit();
	}
	
	function Init(){
	<%
	If Not rsEmailFields.EOF Then
	%>
		document.frm0501.Recipient.focus();
	<%
	End If
	%>
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0501" method="POST" action="m009e0501.asp?intEquip_srv_id=<%=Request("intEquip_srv_id")%>">
<h5>Send Email to Referring Agent</h5>
<hr>
<%
If rsEmailFields.EOF Then
%>
<i>Function not available for this equipment service record.</i>
<%
Else
%>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Recipient:</td>
		<td><input type="text" name="Recipient" value="<%=Trim(rsEmailFields.Fields.Item("chvContact_Email").Value)%>" tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td>CC:</td>
		<td><input type="text" name="CC" tabindex="2"></td>
	</tr>
	<tr>
		<td>Sender:</td>
		<td><input type="text" name="Sender" value="<%=rsSender.Fields.Item("chvEmail").Value%>" tabindex="3"></td>
	</tr>
	<tr>
		<td>Subject:</td>
		<td><input type="text" name="Subject" value="<%=(UserName)%>: Repair" size="75" tabindex="4">
	</tr>
<tr>		
<td valign="top">Message:</td>
<td><textarea name="Message" cols="75" rows="15" tabindex="5" accesskey="L">
Hi <%=(rsEmailFields.Fields.Item("chvContactFst_Name").Value)%>,
Please be advised we have completed the equipment repair for <%=UserName%>.
<%
Select Case rsShippingMethod.Fields.Item("insShip_Method_id").Value
'Loomis
case "4":
%>		
We will be shipping the equipment from our warehouse on <%=rsEmailFields.Fields.Item("dtsDlvy_date").Value%>.
<%
'Dynamex
case "9":
%>
We will be shipping the equipment from our warehouse on <%=rsEmailFields.Fields.Item("dtsDlvy_date").Value%>.
<%
'Case Manager Delivery
case "1":
%>
<%=rsEmailFields.Fields.Item("chvCaseManager").Value%> is delivering the equipment on <%=rsEmailFields.Fields.Item("dtsDlvy_date").Value%> so that the equipment can be set up and training on adaptive software conducted.
<%
End Select
%>
If you have any questions, please do not hesitate to call me at (604)959-8188.
		
Sales Geek

The Geek
Sirius Innovations Inc.
P.O. Box 43119 Richmond CTR PO,Richmond, B.C. V6V 2W4
O: (604)959-8188
F: (604)959-3169

</textarea></td>
</tr>
</table>
<input type="submit" value="Send" class="btnstyle" tabindex="6">
<%
End If
%>
<input type="hidden" name="MM_send" value="1">
</form>
</body>
</html>
<%
set rsSender = Nothing
Set rsEmailFields = Nothing
Set rsShippingMethod = Nothing
%>
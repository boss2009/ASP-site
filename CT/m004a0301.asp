<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_Insert"))=="true") {
	var rsMailingList = Server.CreateObject("ADODB.Recordset");
	rsMailingList.ActiveConnection = MM_cnnASP02_STRING;
	rsMailingList.Source="{call dbo.cp_ctc_mail_list(0,"+Request.Form("intContact_id")+","+Request.Form("MailingList")+",0,'A',0)}";
	rsMailingList.CursorType = 0;
	rsMailingList.CursorLocation = 2;
	rsMailingList.LockType = 3;
	rsMailingList.Open();
	Response.Redirect("InsertSuccessful.html");
}

var rsMailingList = Server.CreateObject("ADODB.Recordset");
rsMailingList.ActiveConnection = MM_cnnASP02_STRING;
rsMailingList.Source = "{call dbo.cp_mail_list(0,'',1,0,'Q',0)}";
rsMailingList.CursorType = 0;
rsMailingList.CursorLocation = 2;
rsMailingList.LockType = 3;
rsMailingList.Open();
%>
<html>
<head>
	<title>Subscribe Mailing List</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
		   	case 76 :
				//alert("L");
				self.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		document.frm0301.submit();
	}	
	</script>
</head>
<body onLoad="document.frm0301.MailingList.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0301">
<h5>Subscribe Mailing List</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><select name="MailingList" tabindex="1" accesskey="F">
		<%
		while (!rsMailingList.EOF) {			
		%>
			<option value="<%=rsMailingList.Fields.Item("insMail_list_id").Value%>"><%=rsMailingList.Fields.Item("chvName").Value%>
		<%
			rsMailingList.MoveNext();
		}
		%>
        </select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="3" onClick="self.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="intContact_id" value="<%=Request.QueryString("intContact_id")%>">
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>
<%
rsMailingList.Close();
%>
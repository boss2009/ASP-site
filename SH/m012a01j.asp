<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) != "undefined") {
	var rsDesktop = Server.CreateObject("ADODB.Recordset");
	rsDesktop.ActiveConnection = MM_cnnASP02_STRING;
	rsDesktop.Source = "{call dbo.cp_Insert_DeskTop(" + Session("insStaff_id") + "," + Request.Form("insSchool_id") + ",27,'" + CurrentDate() + "',0)}";
	rsDesktop.CursorType = 0;
	rsDesktop.CursorLocation = 2;
	rsDesktop.LockType = 3;
	rsDesktop.Open();
	Response.Redirect("InsertSuccessful.html");	
}

var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_school2("+Request.QueryString("insSchool_id")+",'',0,0,0,0,0,0,0,'',1,'Q',0)}"
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();	

var rsModule = Server.CreateObject("ADODB.Recordset");
rsModule.ActiveConnection = MM_cnnASP02_STRING;
rsModule.Source = "{call dbo.cp_ASP_Lkup(700)}";
rsModule.CursorType = 0;
rsModule.CursorLocation = 2;
rsModule.LockType = 3;
rsModule.Open();
%>
<html>
<head>
	<title>Save to Desktop</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
		case 83 :
			//alert("S");
			document.frm12j.submit();
			break;
	   	case 76 :
			//alert("L");
			window.close();
			break;
		}
	}
	</script>	
</head>
<body onLoad="window.focus();">
<form name="frm12j" method="POST" action="<%=MM_editAction%>">
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Institution Name:</td>
		<td><input type="text" name="Institution" value="<%=rsInstitution.Fields.Item("chvSchool_Name").Value%>" readonly size="35" tabindex="1" accesskey="F"></td>
    </tr>	
    <tr> 
		<td>Save To:</td>
		<td><select name="SaveTo" tabindex="2">
			<% 
			while (!rsModule.EOF) {
			%>
				<option value="<%=(rsModule.Fields.Item("insId").Value)%>" <%=((rsModule.Fields.Item("insId").Value == 27)?"SELECTED":"")%> ><%=(rsModule.Fields.Item("ncvMODname").Value)%>
			<%
				rsModule.MoveNext();
			}
			%>
		</select></td>
    </tr>
    <tr> 
		<td>Date Copied:</td>
		<td>
			<input type="text" name="DateCopied" value="<%=CurrentDate()%>" size="10" readonly tabindex="3" accesskey="L">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>						
		</td>
    </tr>
    <tr> 
		<td colspan="2" align="center">
			<input type="submit" value="Save to Desktop" class="btnstyle">
			<input type="button" value="Cancel" onClick="self.close();" class="btnstyle">
		</td>
    </tr>
</table>
<input type="hidden" name="insSchool_id" value="<%=rsInstitution.Fields.Item("insSchool_id").Value%>">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsInstitution.Close();
rsModule.Close();
%>
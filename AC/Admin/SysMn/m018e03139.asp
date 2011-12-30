<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var rsQualificationSource = Server.CreateObject("ADODB.Recordset");
	rsQualificationSource.ActiveConnection = MM_cnnASP02_STRING;
	rsQualificationSource.Source = "{call dbo.cp_grant_qlf_src("+Request.Form("MM_recordId")+",'"+Description+"',0,'E',0)}";
	rsQualificationSource.CursorType = 0;
	rsQualificationSource.CursorLocation = 2;
	rsQualificationSource.LockType = 3;
	rsQualificationSource.Open();
	Response.Redirect("m018q03139.asp");
}

var rsQualificationSource = Server.CreateObject("ADODB.Recordset");
rsQualificationSource.ActiveConnection = MM_cnnASP02_STRING;
rsQualificationSource.Source = "{call dbo.cp_grant_qlf_src("+ Request.QueryString("insGrant_Qlf_Src") + ",'',1,'Q',0)}";
rsQualificationSource.CursorType = 0;
rsQualificationSource.CursorLocation = 2;
rsQualificationSource.LockType = 3;
rsQualificationSource.Open();
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>Update Grant Qualification Source Lookup</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm03139.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (Trim(document.frm03139.Description.value)==""){
			alert("Enter Description.");
			document.frm03139.Description.focus();
			return ;		
		}
		document.frm03139.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03139.Description.focus();">
<form name="frm03139" method="POST" action="<%=MM_editAction%>">
<h5>Update Grant Qualification Source Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=Trim(rsQualificationSource.Fields.Item("chvGrant_Qlf_Src").Value)%>" maxlength="40" size="20" tabindex="1" accesskey="F" ></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsQualificationSource.Fields.Item("insGrant_Qlf_Src").Value %>">
</form>
</body>
</html>
<%
rsQualificationSource.Close();
%>
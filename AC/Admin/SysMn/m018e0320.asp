<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_edit"))=="true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var IsActive = ((Request.Form("IsActive")=="1")?"1":"0");
	var rsProgramType = Server.CreateObject("ADODB.Recordset");
	rsProgramType.ActiveConnection = MM_cnnASP02_STRING;
	rsProgramType.Source = "{call dbo.cp_program_type2("+Request.QueryString("insProg_type_id")+",'" + Description + "'," + IsActive + ",0,'E',0)}";
	rsProgramType.CursorType = 0;
	rsProgramType.CursorLocation = 2;
	rsProgramType.LockType = 3;
	rsProgramType.Open();
	Response.Redirect("m018q0320.asp");
}
var rsProgramType = Server.CreateObject("ADODB.Recordset");
rsProgramType.ActiveConnection = MM_cnnASP02_STRING;
rsProgramType.Source = "{call dbo.cp_program_type2("+Request.QueryString("insProg_type_id")+",'',0,1,'Q',0)}";
rsProgramType.CursorType = 0;
rsProgramType.CursorLocation = 2;
rsProgramType.LockType = 3;
rsProgramType.Open();
%>
<html>
<head>
	<title>Update Program Type Lookup</title>
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
				document.frm0320.reset();
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
		if (Trim(document.frm0320.Description.value)==""){
			alert("Enter Description.");
			document.frm0320.Description.focus();
			return ;		
		}
		document.frm0320.submit();
	}
	</script>
</head>
<body onLoad="document.frm0320.Description.focus();">
<form name="frm0320" method="POST" action="<%=MM_editAction%>">
<h5>Update Program Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" maxlength="50" size="30" tabindex="1" accesskey="F" value="<%=(rsProgramType.Fields.Item("chvname").Value)%>"></td>
	</tr>
	<tr> 
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" value="1" tabindex="2" accesskey="L" <%if (rsProgramType.Fields.Item("bitIsAdd").Value=="1") Response.Write("CHECKED");%> class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="5" onClick="history.back()" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_edit" value="true">
</form>
</body>
</html>
<%
rsProgramType.Close();
%>
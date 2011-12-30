<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update"))=="true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var IsActive = ((Request.Form("IsActive")=="on")?"1":"0");
	var rsWorkOrder = Server.CreateObject("ADODB.Recordset");
	rsWorkOrder.ActiveConnection = MM_cnnASP02_STRING;
	rsWorkOrder.Source = "{call dbo.cp_Work_Order("+Request.QueryString("insWork_order_id")+",'" + Request.Form("WorkOrderNumber") + "','" + Request.Form("Description") + "'," + IsActive + ",0,'E',0)}";
	rsWorkOrder.CursorType = 0;
	rsWorkOrder.CursorLocation = 2;
	rsWorkOrder.LockType = 3;
	rsWorkOrder.Open();
	Response.Redirect("m018q0359.asp");
}

var rsWorkOrder = Server.CreateObject("ADODB.Recordset");
rsWorkOrder.ActiveConnection = MM_cnnASP02_STRING;
rsWorkOrder.Source = "{call dbo.cp_Work_Order("+Request.QueryString("insWork_order_id")+",'','',0,1,'Q',0)}";
rsWorkOrder.CursorType = 0;
rsWorkOrder.CursorLocation = 2;
rsWorkOrder.LockType = 3;
rsWorkOrder.Open();
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
	<title>Update Work Order Lookup</title>
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
				document.frm0359.reset();
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
		if (Trim(document.frm0359.Description.value)==""){
			alert("Enter Description.");
			document.frm0359.Description.focus();
			return ;		
		}
		document.frm0359.submit();
	}
	</script>
</head>
<body onLoad="document.frm0359.Description.focus();">
<form name="frm0359" method="POST" action="<%=MM_updateAction%>">
<h5>Work Order Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="30" tabindex="1" accesskey="F" value="<%=Trim(rsWorkOrder.Fields.Item("chvWork_order_Desc").Value)%>"></td>
	</tr>
	<tr> 
		<td>Work Order Number:</td>
		<td><input type="text" name="WorkOrderNumber" maxlength="40" size="20" tabindex="2" value="<%=Trim(rsWorkOrder.Fields.Item("chvWork_order_no").Value)%>"></td>
	</tr>
	<tr> 
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" tabindex="3" accesskey="L" <%=((rsWorkOrder.Fields.Item("bitis_Active").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="4" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="6" onClick="history.back()" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsWorkOrder.Close();
%>
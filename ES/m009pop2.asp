<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update"))=="true"){
	var rsUpdateBoxes = Server.CreateObject("ADODB.Recordset");
	rsUpdateBoxes.ActiveConnection = MM_cnnASP02_STRING;
	rsUpdateBoxes.CursorType = 0;
	rsUpdateBoxes.CursorLocation = 2;
	rsUpdateBoxes.LockType = 3;
	if (Request.Form("BoxCount") > 1) {
		for (i = 1; i <= Request.Form("BoxCount"); i++){
			if (Request.Form("BoxWeight")(i) > 0) {
				rsUpdateBoxes.Source = "{call dbo.cp_eqpsrv_ship_box("+Request.Form("ID")(i)+","+Request.QueryString("intSrv_dtl_id")+",0,"+Request.Form("BoxWeight")(i)+","+Request.QueryString("intEquip_Srv_id")+",0,0,'E',0)}";
				rsUpdateBoxes.Open();			
			} else {
				rsUpdateBoxes.Source = "{call dbo.cp_eqpsrv_ship_box("+Request.Form("ID")(i)+","+Request.QueryString("intSrv_dtl_id")+",0,0,"+Request.QueryString("intEquip_Srv_id")+",0,0,'D',0)}";	
				rsUpdateBoxes.Open();						
			}
		}
	}
	if (Request.Form("BoxCount") == 1) {
		if (Request.Form("BoxWeight") > 0) {	
			rsUpdateBoxes.Source = "{call dbo.cp_eqpsrv_ship_box("+Request.Form("ID")+","+Request.QueryString("intSrv_dtl_id")+",0,"+Request.Form("BoxWeight")+","+Request.QueryString("intEquip_Srv_id")+",0,0,'E',0)}";
			rsUpdateBoxes.Open();			
		} else {
			rsUpdateBoxes.Source = "{call dbo.cp_eqpsrv_ship_box("+Request.Form("ID")+","+Request.QueryString("intSrv_dtl_id")+",0,0,"+Request.QueryString("intEquip_Srv_id")+",0,0,'D',0)}";	
			rsUpdateBoxes.Open();						
		}
	}
	
	if (Request.Form("NewBoxWeight") > 0) {	
		rsUpdateBoxes.Source = "{call dbo.cp_eqpsrv_ship_box(0,"+Request.QueryString("intSrv_dtl_id")+",1,"+Request.Form("NewBoxWeight")+","+Request.QueryString("intEquip_Srv_id")+",0,0,'A',0)}";
		rsUpdateBoxes.Open();				
	}	
}

var rsBoxes = Server.CreateObject("ADODB.Recordset");
rsBoxes.ActiveConnection = MM_cnnASP02_STRING;
rsBoxes.Source = "{call dbo.cp_eqpsrv_ship_box(0,"+Request.QueryString("intSrv_dtl_id")+",0,0,"+Request.QueryString("intEquip_Srv_id")+",0,0,'Q',0)}";
rsBoxes.CursorType = 0;
rsBoxes.CursorLocation = 2;
rsBoxes.LockType = 3;
rsBoxes.Open();
%>
<html>
<head>
	<title>Shipping Boxes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function Save(){
		document.frmpop2.submit();
	}
		
	function Init(){
		document.frmpop2.NewBoxWeight.focus();
	}
	
	function Close(){
		var TotalWeight = 0;
		if (document.frmpop2.BoxCount.value == 1) TotalWeight = document.frmpop2.BoxWeight.value;
		if (document.frmpop2.BoxCount.value > 1) {
			for (i = 0; i < document.frmpop2.BoxCount.value; i++) {
				TotalWeight = TotalWeight + Number(document.frmpop2.BoxWeight[i].value);
			}
		}
		if (!opener.closed) {		
			opener.frm0302.NumberOfBoxes.value=document.frmpop2.BoxCount.value;
			opener.frm0302.TotalWeight.value=TotalWeight;
		}
		window.close();
	}
	</script>
</head>
<body onLoad="Init();">
<form name="frmpop2" method="POST" action="<%=MM_editAction%>">
<h5>Shipping Boxes</h5>
<hr>
<table cellspacing="1" cellpadding="1">
<% 
var row = 0;
while (!rsBoxes.EOF) { 
	row++;
%>
    <tr> 
		<td><%=row%>.</td>
		<td><input type="text" name="BoxWeight" value="<%=(rsBoxes.Fields.Item("insBox_Wgt").Value)%>" size="3" onKeypress="AllowNumericOnly();"> LB (<%=rsBoxes.Fields.Item("insBox_Wgt").Value*0.454%>)kg</td>
    </tr>
	<input type="hidden" name="ID" value="<%=rsBoxes.Fields.Item("insSB_id").Value%>">	
<%
	rsBoxes.MoveNext();
}
%>
</table>
<hr>
Add Box Weight: <input type="text" name="NewBoxWeight" size="4" onKeypress="AllowNumericOnly();"> LB
<input type="button" value="Save" onClick="Save();" class="btnstyle">&nbsp;
<input type="button" value="Close" onClick="Close();" class="btnstyle">
<input type="hidden" name="BoxCount" value="<%=row%>">
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsBoxes.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsEquipmentService = Server.CreateObject("ADODB.Recordset");
rsEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentService.Source = "{call dbo.cp_get_eqp_srv("+ Request.QueryString("intEquip_Srv_id") + ",0,0,'',1,'Q',0)}";
rsEquipmentService.CursorType = 0;
rsEquipmentService.CursorLocation = 2;
rsEquipmentService.LockType = 3;
rsEquipmentService.Open();
%>
<html>
<head>
	<title>Current Inventory Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script langauge="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	
	function ShowFundingSource(id) {
		var temp = window.showModalDialog("m009pop.asp?intEquip_Set_id="+id,"","dialogHeight: 350px; dialogWidth: 400px; dialogTop: px; dialogLeft: px; edge: Sunken; center: Yes; help: No; resizable: No; status: No;");	
	}
	</script>
</head>
<body>
<h5>Current Inventory Information</h5>
<table cellpadding="2" cellspacing="3">
	<tr> 
		<td nowrap>Inventory Name:</td>
		<td nowrap><a href="javascript: openWindow('../IV/m003FS3.asp?intEquip_Set_id=<%=rsEquipmentService.Fields.Item("intEquip_Set_id").Value%>&intBar_Code_no=<%=rsEquipmentService.Fields.Item("intEquip_Set_id").Value%>','');"><%=(rsEquipmentService.Fields.Item("chvInventory_Name").Value)%></a></td>	
	</tr>
	<tr> 
		<td nowrap>Inventory Status:</td>
		<td nowrap><%=(rsEquipmentService.Fields.Item("chvCurrent_Status").Value)%></td>
	</tr>
	<tr>		
		<td nowrap>Warranty Start Date:</td>
		<td nowrap><%=FilterDate(rsEquipmentService.Fields.Item("dtsRec_Date").Value)%><%if ((FilterDate(rsEquipmentService.Fields.Item("dtsRec_Date").Value)!="") && (rsEquipmentService.Fields.Item("dtsRec_Date").Value!=null)) {%><span style="font-size: 7pt">(mm/dd/yyyy)</span><%}%></td>
	</tr>
	<tr>
		<td nowrap>Parts Warranty:</td>
		<td nowrap><%=(rsEquipmentService.Fields.Item("chvPartsWLen").Value)%></td>
	</tr>
    <tr> 
		<td nowrap>Serial Number:</td>
		<td nowrap><%=(rsEquipmentService.Fields.Item("chvSerial_Number").Value)%></td>
	</tr>
	<tr>
		<td nowrap>Labour Warranty:</td>
		<td nowrap><%=(rsEquipmentService.Fields.Item("chvLaborWLen").Value)%></td>
	</tr>
	<tr> 
		<td nowrap>Model Number:</td>
		<td nowrap><%=(rsEquipmentService.Fields.Item("chvModel_Number").Value)%></td>
	</tr>
	<tr>
		<td nowrap>PR Number:</td>
		<td nowrap><%=(rsEquipmentService.Fields.Item("intRequisition_no").Value)%></td>				
    </tr>
    <tr> 
		<td nowrap>Vendor:</td>
		<td nowrap><%=(rsEquipmentService.Fields.Item("chvVendor").Value)%></td>
	</tr>
<!--	
	<tr> 
		<td nowrap>Funding Source:</td>
		<td nowrap><input type="button" value="View" onClick="ShowFundingSource(<%=(rsEquipmentService.Fields.Item("intEquip_Set_id").Value)%>);" tabindex="4" class="btnstyle"></td>
	</tr>
-->
</table>
</body>
</html>
<%
rsEquipmentService.Close();
%>
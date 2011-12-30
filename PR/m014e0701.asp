<!--------------------------------------------------------------------------
* File Name: m014e0701.asp
* Title: Create Inventory
* Main SP: cp_get_purchase_requisition, cp_verify_bcdnum, cp_insert_eqcls_inventory_2
* Description: The page loops through all the inventories, checks the InventoryID
* to make sure a duplicate does not exist on the database.  If ID is unique,
* the inventory is created, else all the inventories entered are rolled back.
* Author: T.H
--------------------------------------------------------------------------->
<%@TRANSACTION=Required language="VBScript"%>
<!--#include file="../inc/VBLogin.inc"-->
<%
if Request.Form("State")="Save" then
	dim InventoryName	
	dim rsRequisition	
	set rsRequisition = Server.CreateObject("ADODB.Recordset")
	rsRequisition.ActiveConnection = MM_cnnASP02_STRING
	rsRequisition.Source = "{call dbo.cp_Get_Purchase_Requisition(1,0,'',1,"&Request.Form("insPurchase_Req_id")&",0)}"
	rsRequisition.CursorType = 0
	rsRequisition.CursorLocation = 2
	rsRequisition.LockType = 3
	rsRequisition.Open()
	dim OrderedDate
	OrderedDate = rsRequisition.Fields.Item("dtsDate_Ordered").Value
	rsRequisition.Close()
	set rsRequistion= nothing	
	if CInt(Request.Form("Quantity")) = 1 then
		Dim rsOneBarcode
		set rsOneBarcode = Server.CreateObject("ADODB.Recordset")
		rsOneBarcode.ActiveConnection = MM_cnnASP02_STRING
		rsOneBarcode.CursorType = 0
		rsOneBarcode.CursorLocation = 2
		rsOneBarcode.LockType = 3	
		rsOneBarcode.Source = "{call dbo.cp_Verify_BCdnum("&Request.Form("InventoryID")&",0)}"	
		rsOneBarcode.Open()		
		if Not rsOneBarcode.EOF then
			ObjectContext.SetAbort
		end if	
		set rsOneBarcode = nothing
		Dim rsNewInventory 
		set rsNewInventory = Server.CreateObject("ADODB.Recordset")		
		rsNewInventory.ActiveConnection = MM_cnnASP02_STRING		
		rsNewInventory.CursorType = 0
		rsNewInventory.CursorLocation = 2
		rsNewInventory.LockType = 3		
		InventoryName = Replace(Request("InventoryName"), "'", "''")   										
		rsNewInventory.Source = "{call dbo.cp_Insert_EqCls_Inventory_02("&Request.Form("InventoryClass")&","&Request.Form("InventoryID")&",'"&Request.Form("SerialNumber")& "'," &Request.Form("insPurchase_Req_id")& ","&Request.Form("EquipmentCost")&",'" &InventoryName& "',1,1,'" &OrderedDate& "','" &Request.Form("DateReceived")& "','',0," & Session("insStaff_id")&",0)}"
		rsNewInventory.Open()
		set rsNewInventory = nothing	
	else
		for i = 1 to CInt(Request.Form("Quantity"))
			Dim rsBarcodes
			set rsBarcodes = Server.CreateObject("ADODB.Recordset")
			rsBarcodes.ActiveConnection = MM_cnnASP02_STRING
			rsBarcodes.CursorType = 0
			rsBarcodes.CursorLocation = 2
			rsBarcodes.LockType = 3		
			rsBarcodes.Source = "{call dbo.cp_Verify_BCdnum("&Request.Form("InventoryID")(i)&",0)}"	
			rsBarcodes.Open()		
			if Not rsBarcodes.EOF then
				ObjectContext.SetAbort
			end if
			set rsBarcodes = nothing
			Dim rsNewInventories
			set rsNewInventories = Server.CreateObject("ADODB.Recordset")		
			rsNewInventories.ActiveConnection = MM_cnnASP02_STRING		
			rsNewInventories.CursorType = 0
			rsNewInventories.CursorLocation = 2
			rsNewInventories.LockType = 3
			InventoryName = Replace(Request("InventoryName")(i), "'", "''")   								
			rsNewInventories.Source = "{call dbo.cp_Insert_EqCls_Inventory_02("&Request.Form("InventoryClass")&","&Request.Form("InventoryID")(i)&",'"&Request.Form("SerialNumber")(i)& "'," &Request.Form("insPurchase_Req_id")& ","&Request.Form("EquipmentCost")& ",'" &InventoryName& "',1,1,'" &OrderedDate& "','" &Request.Form("DateReceived")& "','',0," & Session("insStaff_id")&",0)}"
			rsNewInventories.Open()
			set rsNewInventories = nothing
		next 
	end if
	Response.Redirect("InsertSuccessful.html")
end if

Sub OnTransactionAbort()
	Response.Redirect("Failed.asp")
End Sub

Dim rsInventoryRequested
set rsInventoryRequested = Server.CreateObject("ADODB.Recordset")
rsInventoryRequested.ActiveConnection = MM_cnnASP02_STRING
rsInventoryRequested.Source = "{call dbo.cp_Purchase_Requisition_Requested(0,"&Request("insPurchase_Req_id")&",0,0,0,'',0.0,'01/01/1999',0,0,0,'Q',0)}"
rsInventoryRequested.CursorType = 0
rsInventoryRequested.CursorLocation = 2
rsInventoryRequested.LockType = 3
rsInventoryRequested.Open()

Dim rsInventoryReceived 
set rsInventoryReceived = Server.CreateObject("ADODB.Recordset")
rsInventoryReceived.ActiveConnection = MM_cnnASP02_STRING
rsInventoryReceived.Source = "{call dbo.cp_Purchase_Requisition_Received("&Request("insPurchase_Req_id")&",0,0,'',0,'',0,'Q',0)}"
rsInventoryReceived.CursorType = 0
rsInventoryReceived.CursorLocation = 2
rsInventoryReceived.LockType = 3
rsInventoryReceived.Open()

Dim i
i = 0
While Not rsInventoryReceived.EOF
	i = i + rsInventoryReceived.Fields.Item("intQuantity_Received").Value
	rsInventoryReceived.MoveNext()
Wend
%>
<html>
<head>
	<title>Create Inventory</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				if (!document.frm0701.btnSave.disabled) Save();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
			case 71:
				//alert("G");
				Generate();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Generate() {
		if (document.frm0701.Quantity.value=="") document.frm0701.Quantity.value="0";
		if ((!IsID(document.frm0701.StartingID.value)) || (document.frm0701.StartingID.value=="")) {
			alert("Invalid Starting ID.");
			document.frm0701.StartingID.focus();
			return ;
		}
		var qty = new Number(document.frm0701.Quantity.value);
		if (document.frm0701.InventoryClass.length > 1){
			var maxqty = new Number(document.frm0701.MaximumQuantity[document.frm0701.InventoryClass.selectedIndex].value);		
			if (qty > maxqty) {
				alert("Quantity cannot exceed "+ document.frm0701.MaximumQuantity[document.frm0701.InventoryClass.selectedIndex].value+".");
				document.frm0701.Quantity.focus();
				return;
			}		
		} else {
			var maxqty = new Number(document.frm0701.MaximumQuantity.value);
			if (qty > maxqty) {
				alert("Quantity cannot exceed "+ document.frm0701.MaximumQuantity.value +".");
				document.frm0701.Quantity.focus();
				return;
			}
		}
		document.frm0701.State.value="Generate";
		document.frm0701.submit();
	}
	
	function Save(){
		if (!CheckDate(document.frm0701.DateReceived.value)) {
			alert("Invalid received date.");
			document.frm0701.DateReceived.focus();
			return;
		}
		document.frm0701.State.value="Save";
		document.frm0701.submit();
	}
		
	function ChangeQuantityCost(){
		if (document.frm0701.InventoryClass.length > 1){
			var maxqty = new Number(document.frm0701.MaximumQuantity[document.frm0701.InventoryClass.selectedIndex].value);		
			var cost = new Number(document.frm0701.ListUnitCost[document.frm0701.InventoryClass.selectedIndex].value);					
			document.frm0701.Quantity.value=maxqty;
			document.frm0701.UnitCost.value=cost;
			document.frm0701.EquipmentCost.value=cost;			
			return;
		} else {
			var maxqty = new Number(document.frm0701.MaximumQuantity.value);
			var cost = new Number(document.frm0701.ListUnitCost.value);
			document.frm0701.Quantity.value=maxqty;
			document.frm0701.UnitCost.value=cost;
			document.frm0701.EquipmentCost.value=cost;						
			return;
		}		
	}	
	
	function Init(){		
<%
if Request.Form("State")="Generate" then			
	for j = 0 to Request.Form("Quantity")-1
%>
		if (String(document.frm0701.InventoryName.length)=="undefined") {
			document.frm0701.InventoryName.value=document.frm0701.InventoryClass.options[document.frm0701.InventoryClass.selectedIndex].text;
		} else {			
			document.frm0701.InventoryName[<%=j%>].value=document.frm0701.InventoryClass.options[document.frm0701.InventoryClass.selectedIndex].text;
		}
<%
	next
%>
		document.frm0701.State.value="Generate";
<%
end if
%>
		if ((document.frm0701.State.value=="Generate") && (document.frm0701.Duplicate.value=="false")) {
			document.frm0701.btnSave.disabled = false;
		} else {
			document.frm0701.btnSave.disabled = true;
		}
<%
	if Not Request.Form("State")="Generate" then
%>			
		ChangeQuantityCost();
<%
	end if
%>
		document.frm0701.InventoryClass.focus();
	}
	</script>
</head>
<body onLoad="<%If i > 0 Then Response.Write("Init();")%>">
<form name="frm0701" method=POST action="m014e0701.asp">
<h5>Create Inventory</h5>
<%
If i > 0 Then
%>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Inventory Class:</td>
		<td><select name="InventoryClass" tabindex="1" accesskey="F" style="width: 400px" onChange="ChangeQuantityCost();">
			<%
			while Not rsInventoryRequested.EOF
			%>
				<option value="<%=rsInventoryRequested.Fields.Item("insClass_bundle_id").Value%>" <%if CInt(rsInventoryRequested.Fields.Item("insClass_bundle_id").Value)=CInt(Request.Form("InventoryClass")) then Response.Write("SELECTED") end if%>><%=rsInventoryRequested.Fields.Item("chvClass_Bundle_Name").Value%>
			<%
				rsInventoryRequested.MoveNext			
			wend	
			rsInventoryRequested.MoveFirst	
			%>
		</select><%=Request.Form("InventoryClass")%></td>
	</tr>
	<tr>
		<td>Unit Cost:</td>
		<td>$<input type="text" name="UnitCost" size="10" tabindex="2" readonly value="<%=Request.Form("UnitCost")%>"></td>
	</tr>
	<tr>
		<td>Equipment Cost:</td>
		<td>$<input type="text" name="EquipmentCost" size="10" onKeypress="AllowNumericOnly();" tabindex="3" value="<%=Request.Form("EquipmentCost")%>"></td>
	</tr>	
	<tr>
		<td>Date Received:</td>
		<td>
			<input type="text" name="DateReceived" size="11" maxlength="10"  value="<%=Date()%>" tabindex="4" onChange="FormatDate(this)">
			<span style="size: 8pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td>Starting ID:</td>
		<td><input type="text" name="StartingID" tabindex="5"  size="14" onKeypress="AllowNumericOnly();" value="<%=Request.Form("StartingID")%>"></td>	
	</tr>
	<tr>
		<td>Quantity:</td>
		<td><input type="text" name="Quantity" size="3"  onKeypress="AllowNumericOnly();" tabindex="6" value="<%=Request.Form("Quantity")%>"></td>
	</tr>
	<tr>
		<td>Increment:</td>
		<td><select name="Increment" tabindex="7">
				<option value="1" <%if Request.Form("Increment")="1" then Response.Write("SELECTED") end if%>>1
				<option value="2" <%if Request.Form("Increment")="2" then Response.Write("SELECTED") end if%>>2
				<option value="3" <%if Request.Form("Increment")="3" then Response.Write("SELECTED") end if%>>3
				<option value="4" <%if Request.Form("Increment")="4" then Response.Write("SELECTED") end if%>>4
				<option value="5" <%if Request.Form("Increment")="5" then Response.Write("SELECTED") end if%>>5
		</select>
		<input type="button" value="Generate" onClick="Generate();" tabindex="8" class="btnstyle"></td>
	</tr>
</table>
<br>
<div class="BrowsePanel" style="width: 520px; height: 180px; top: 220px"> 
<table width="100%" cellpadding="0" cellspacing="1">
<%
Dim duplicate 
duplicate = false 
if Request.Form("State")="Generate" then
%>
      <tr> 
        <td class="headrow" nowrap>Inventory ID</td>
        <td class="headrow" nowrap>Serial Number</td>
        <td class="headrow" nowrap>Inventory Name</td>
      </tr>
      <%
Dim start
Dim rsBarcode
set rsBarcode = Server.CreateObject("ADODB.Recordset")
rsBarcode.ActiveConnection = MM_cnnASP02_STRING
rsBarcode.CursorType = 0
rsBarcode.CursorLocation = 2
rsBarcode.LockType = 3

start = Request.Form("StartingID")
for i = 1 to CInt(Request.Form("Quantity"))
	rsBarcode.Source = "{call dbo.cp_Verify_BCdnum("&start&",0)}"	
	rsBarcode.Open()		
	if Not rsBarcode.EOF then
		duplicate = true
	end if
%>
      <tr> 
        <td><input type="text" name="InventoryID" value="<%=start%>" size="14" readonly style="<%if Not rsBarcode.EOF then%>color: red<%end if%>"></td>
        <td><input type="text" name="SerialNumber" size="14"></td>
        <td><input type="text" name="InventoryName" size="47"></td>
      </tr>
      <%
	start = start + CInt(Request.Form("Increment"))
	rsBarcode.Close()
Next
end if
%>
</table>
</div>
<div style="position: absolute; top: 410px">
<%
If duplicate Then
%>
<i>* IDs in <font color=red>red</font> already exist in the database.</i>
<%
End If
%>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" name="btnSave" value="Save" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="self.close();" class="btnstyle"></td>
    </tr>
</table>
</div>
<input type="hidden" name="State" value="Start">
<input type="hidden" name="insPurchase_Req_id" value="<%=Request("insPurchase_Req_id")%>">
<%
If Request.Form("State") = "Generate" AND duplicate Then
%>
<input type="hidden" name="Duplicate" value="true">
<%
Else 
%>
<input type="hidden" name="Duplicate" value="false">
<%
End If
rsInventoryRequested.MoveFirst
while Not rsInventoryRequested.EOF 
%>
<input type="hidden" name="MaximumQuantity" value="<%=rsInventoryRequested.Fields.Item("insPR_request_Qty_Ordered").Value%>">
<input type="hidden" name="ListUnitCost" value="<%=rsInventoryRequested.Fields.Item("fltPR_request_List_Unit_Cost").Value%>">
<%
	rsInventoryRequested.MoveNext
WEnd
%>	
</form>
<%
Else
%>
<i>Receive inventory quantity in Received Page before creating inventory records.</i>
<%
End If
%>
</body>
</html>
<%
set rsInventoryRequested = Nothing
%>
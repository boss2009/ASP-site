<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.Form("MM_Insert"))=="true") {
	var cmdCheckBarcode = Server.CreateObject("ADODB.Command");
	cmdCheckBarcode.ActiveConnection = MM_cnnASP02_STRING;
	cmdCheckBarcode.CommandText = "dbo.cp_Ivtry_Barcodenum";
	cmdCheckBarcode.CommandType = 4;
	cmdCheckBarcode.CommandTimeout = 0;
	cmdCheckBarcode.Prepared = true;
	cmdCheckBarcode.Parameters.Append(cmdCheckBarcode.CreateParameter("RETURN_VALUE", 3, 4));
	cmdCheckBarcode.Parameters.Append(cmdCheckBarcode.CreateParameter("@intBar_Code_no", 3, 1,8,Request.Form("InventoryID")));
	cmdCheckBarcode.Parameters.Append(cmdCheckBarcode.CreateParameter("@intRecCnt", 3, 2));
	cmdCheckBarcode.Execute();
	
	if (cmdCheckBarcode.Parameters.Item("@intRecCnt").Value==0) {
		var InventoryName = String(Request.Form("InventoryName")).replace(/'/g, "'");	
		var bitIs_Template = Request.Form("IsTemplate");
		var InventoryCost = ((String(Request.Form("InventoryCost"))=="")?"0":Request.Form("InventoryCost"));
		var PurchaseRequisitionNumber = ((String(Request.Form("PurchaseRequisitionNumber"))=="")?"0":Request.Form("PurchaseRequisitionNumber"));		
		var DateOrdered = ((String(Request.Form("DateOrdered"))=="undefined")?"1/1/1900":Request.Form("DateOrdered"));
		var DateReceived = ((String(Request.Form("DateReceived"))=="undefined")?"1/1/1900":Request.Form("DateReceived"));
		var cmdInsertInventory = Server.CreateObject("ADODB.Command");
		cmdInsertInventory.ActiveConnection = MM_cnnASP02_STRING;
		cmdInsertInventory.CommandText = "dbo.cp_Insert_EqCls_Inventory_02";
		cmdInsertInventory.CommandType = 4;
		cmdInsertInventory.CommandTimeout = 0;
		cmdInsertInventory.Prepared = true;
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("RETURN_VALUE", 3, 4));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@insEquip_Class_id", 2, 1,1,Request.Form("ConcreteClass")));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@intBar_Code_no", 3, 1,1,Request.Form("InventoryID")));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@chvSerial_Number", 200, 1,20,Request.Form("SerialNumber")));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@intRequisition_no", 3, 1,1,PurchaseRequisitionNumber));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@fltPurchase_Cost", 5, 1,1,InventoryCost));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@chvInventory_Name", 200, 1,80,InventoryName));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@insCurrent_Status", 2, 1,1,Request.Form("Status")));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@chrType", 129, 1,1,Request.Form("Type")));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@dtsOrd_Date", 200, 1,30,DateOrdered));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@dtsRec_Date", 200, 1,30,DateReceived));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@chvActiveKey", 200, 1,50,Request.Form("ActivationKey")));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@bitIs_Template", 2, 1,1,bitIs_Template));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@insCreator_user_id", 2, 1,1,Session("insStaff_id")));
		cmdInsertInventory.Parameters.Append(cmdInsertInventory.CreateParameter("@intEquip_Set_id", 3, 2));
		cmdInsertInventory.Execute();
		
		if (String(Request.Form("SaveFields"))=="false") Response.Redirect("m003FS3.asp?intEquip_Set_id="+cmdInsertInventory.Parameters.Item("@intEquip_Set_id").Value+"&intBar_Code_no="+cmdInsertInventory.Parameters.Item("@intEquip_Set_id").Value);
	} else {
		Response.Redirect("DuplicateID.asp");
	}
}

var rsAbstractClass = Server.CreateObject("ADODB.Recordset");
rsAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW(0,'A',0)}";
rsAbstractClass.CursorType = 0;
rsAbstractClass.CursorLocation = 2;
rsAbstractClass.LockType = 3;
rsAbstractClass.Open();

var rsSubAbstractClass__ClassID;

if (String(Request.Form("Initialized")) == "true") {
	var rsSubAbstractClass__ClassID = Request.Form("AbstractClass");
} else {
	if (!rsAbstractClass.EOF) rsSubAbstractClass__ClassID = rsAbstractClass.Fields.Item("insEquip_Class_id").Value;	
}

var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsSubAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW("+rsSubAbstractClass__ClassID+",'S',0)}";
rsSubAbstractClass.CursorType = 0;
rsSubAbstractClass.CursorLocation = 2;
rsSubAbstractClass.LockType = 3;
rsSubAbstractClass.Open();

if ((String(Request.Form("Initialized")) == "true") && (String(Request.Form("CInitialize")) == "false")) {
	var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
	rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
	rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW(" + Request.Form("SubAbstractClass") + ",'C',0)}";
	rsConcreteClass.CursorType = 0;
	rsConcreteClass.CursorLocation = 2;
	rsConcreteClass.LockType = 3;
	rsConcreteClass.Open();	
	
	var ConcreteVendor	= ((Request.Form("ConcreteClass")=="")?rsConcreteClass.Fields.Item("insEquip_Class_id").Value:Request.Form("ConcreteClass"));

	//open rsVendor to get concrete class vendor info
	var rsVendor = Server.CreateObject("ADODB.Recordset");
	rsVendor.ActiveConnection = MM_cnnASP02_STRING;
	rsVendor.Source = "{call dbo.cp_Eqp_Class_LW("+ConcreteVendor+",'C',1)}";
	rsVendor.CursorType = 0;
	rsVendor.CursorLocation = 2;
	rsVendor.LockType = 3;
	rsVendor.Open();	

	var rsVendorName = Server.CreateObject("ADODB.Recordset");
	rsVendorName.ActiveConnection = MM_cnnASP02_STRING;
	rsVendorName.Source = "{call dbo.cp_get_Company_Address("+rsVendor.Fields.Item("insVendor_id").Value+",1)}";
	rsVendorName.CursorType = 0;
	rsVendorName.CursorLocation = 2;
	rsVendorName.LockType = 3;
	rsVendorName.Open();	
	
	var InitVendorInfo = true;
}

var rsWarrantyLength = Server.CreateObject("ADODB.Recordset");
rsWarrantyLength.ActiveConnection = MM_cnnASP02_STRING;
rsWarrantyLength.Source = "{call dbo.cp_ASP_lkup(62)}";
rsWarrantyLength.CursorType = 0;
rsWarrantyLength.CursorLocation = 2;
rsWarrantyLength.LockType = 3;
rsWarrantyLength.Open();

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_ASP_lkup(36)}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();	
%>
<html>
<head>
	<title>New Inventory</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save1(1);
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function SelectClass(){
		document.frm0101.submit();
	}		
	
	function Save1(savetype){	
		if (document.frm0101.ConcreteClass.value==""){
			alert("Select a Concrete Class.");
			document.frm0101.ConcreteClass.focus();
			return ;
		}
		if (Trim(document.frm0101.InventoryName.value)==""){
			alert("Enter Inventory Name.");
			document.frm0101.InventoryName.focus();
			return ;
		}
		if (Trim(document.frm0101.InventoryID.value)==""){
			alert("Enter Inventory ID.");
			document.frm0101.InventoryID.focus();
			return ;
		}		
		var temp = new Number(document.frm0101.InventoryID.value);
		if ((temp > 2000000000) || (temp < 0)){
			alert("Inventory ID is out of acceptable range.");
			document.frm0101.InventoryID.focus();
			return;
		}
		if (!CheckDate(document.frm0101.DateOrdered.value)){
			alert("Invalid Date Ordered.");
			document.frm0101.DateOrdered.focus();
			return ;
		}
		if (!CheckDate(document.frm0101.DateReceived.value)){
			alert("Invalid Date Received.");
			document.frm0101.DateReceived.focus();
			return ;
		}
		if (isNaN(document.frm0101.InventoryCost.value)){
			alert("Invalid Equiment Cost.");
			document.frm0101.InventoryCost.focus();
			return ;
		}
		document.frm0101.MM_Insert.value="true";
		if (savetype==2) document.frm0101.SaveFields.value="true";
		if (savetype==3) document.frm0101.IsTemplate.value="1";
		document.frm0101.submit();
	}
	
	function Init(){
		document.frm0101.AbstractClass.focus();
		if (document.frm0101.ConcreteClass.value!='') { 
			document.frm0101.InventoryName.value=<%if (String(Request.Form("SaveFields"))=="true"){Response.Write("'"+Request.Form("InventoryName")+"';")} else { Response.Write("document.frm0101.ConcreteClass[document.frm0101.ConcreteClass.selectedIndex].text;")}%>
		} else { 
			document.frm0101.InventoryName.value='';
		}
		if ((document.frm0101.SubAbstractClass.value > 0) && (document.frm0101.SubAbstractClass.length==1) && (document.frm0101.ConcreteClass.length==1) && (document.frm0101.ConcreteClass.value == 0)){
			document.frm0101.CInitialize.value='false'		
			document.frm0101.submit();
		}
	<%
	if (String(Request.Form("SaveFields"))=="true"){
	%>
		alert("Previous inventory record successfully created.");
	<%
	}
	%>	
	}
	</script>
</head>
<body onLoad="Init();">
<form name="frm0101" method="POST" action="m003a0101.asp">
<h5>New Inventory</h5>
<i>Use [Save and Retain Fields] button if you are creating more than one record for identical items.</i>
<hr>
<b style="font-size: 11pt; color: #8CAAE6;">Class Information</b>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Abstract Class:</td>
		<td nowrap><select name="AbstractClass" accesskey="F" tabindex="1" onChange="document.frm0101.CInitialize.value='true';SelectClass();" style="width: 200px">
			<% 
			while (!rsAbstractClass.EOF){ 
			%>
				<option value="<%=rsAbstractClass.Fields.Item("insEquip_Class_id").Value%>" <%=((Request.Form("AbstractClass")==rsAbstractClass.Fields.Item("insEquip_Class_id").Value)?" SELECTED":"")%>><%=rsAbstractClass.Fields.Item("chvName").Value%> 
			<%
				rsAbstractClass.MoveNext();
			}
			rsAbstractClass.MoveFirst();
			%>
		</select></td>
		<td nowrap>Subject To:</td>
		<td nowrap><select name="SubjectTo" disabled>
			<option value="0" <%if (InitVendorInfo) {Response.Write((rsVendor.Fields.Item("chvSbjTotax").Value == "0")?"SELECTED":"")}%>>No Tax 
			<option value="1" <%if (InitVendorInfo) {Response.Write((rsVendor.Fields.Item("chvSbjTotax").Value == "1")?"SELECTED":"")}%>>PST 
			<option value="2" <%if (InitVendorInfo) {Response.Write((rsVendor.Fields.Item("chvSbjTotax").Value == "2")?"SELECTED":"")}%>>GST 
			<option value="3" <%if (InitVendorInfo) {Response.Write((rsVendor.Fields.Item("chvSbjTotax").Value == "3")?"SELECTED":"")}%>>Both 
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Sub Abstract Class:</td>
		<td nowrap><select name="SubAbstractClass" tabindex="2" style="width: 200px" onChange="SelectClass();">
			<%
			while (!rsSubAbstractClass.EOF){ 
			%>
				<option value="<%=rsSubAbstractClass.Fields.Item("insEquip_Class_id").Value%>" <%=((Request.Form("SubAbstractClass")==rsSubAbstractClass.Fields.Item("insEquip_Class_id").Value)?" SELECTED":"")%>><%=rsSubAbstractClass.Fields.Item("chvName").Value%> 
			<%
				rsSubAbstractClass.MoveNext();
			}
			rsSubAbstractClass.MoveFirst();
			%>
        </select></td>
		<td nowrap>Parts Warranty:</td>
		<td nowrap><select name="PartsWarrantyLength" disabled>
			<% 
			while (!rsWarrantyLength.EOF) { 
			%>
				<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" <%if (InitVendorInfo) {((rsVendor.Fields.Item("insPartsWLen").Value==rsWarrantyLength.Fields.Item("insWarrenty_id").Value)?Response.Write("SELECTED"):Response.Write(""))}%>><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%> 
			<% 
				rsWarrantyLength.MoveNext();			
			} 
			%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Concrete Class:</td>
		<td nowrap><select name="ConcreteClass" tabindex="3" style="width: 200px" onChange="SelectClass();">
		<%
		var tempName = "";
		if ((String(Request.Form("Initialized")) == "true") && (String(Request.Form("CInitialize")) == "false")) {
		//tempName = rsConcreteClass.Fields.Item("chvName").Value;
			while (!rsConcreteClass.EOF) { 
		%>
				<option value="<%=rsConcreteClass.Fields.Item("insEquip_Class_id").Value%>" <%=((Request.Form("ConcreteClass")==rsConcreteClass.Fields.Item("insEquip_Class_id").Value)?"SELECTED":"")%>><%=rsConcreteClass.Fields.Item("chvName").Value%> 
		<%	
				if (Request.Form("ConcreteClass")==rsConcreteClass.Fields.Item("insEquip_Class_id").Value) tempName = rsConcreteClass.Fields.Item("chvName").Value;						
				rsConcreteClass.MoveNext();
			}
		} else {
		%>
				<option value="">Select Sub Abstract Class 
		<%
		}			
		%>
		</select></td>
		<td nowrap>Labour Warranty:</td>
		<td nowrap><select name="LabourWarrantyLength" disabled>
			<% 
			rsWarrantyLength.MoveFirst();			
			while (!rsWarrantyLength.EOF) { 			
			%>
				<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" <%if (InitVendorInfo) { Response.Write((rsVendor.Fields.Item("insLaborWLen").Value==rsWarrantyLength.Fields.Item("insWarrenty_id").Value)?"SELECTED":"")}%>><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%> 
			<% 
				rsWarrantyLength.MoveNext();
			} 
			%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap>Default Vendor:</td>
		<td nowrap><input type="text" name="DefaultVendor" value="<%=((InitVendorInfo && !rsVendorName.EOF)?rsVendorName.Fields.Item("chvName").Value:"")%>" size="30" readonly></td>
		<td nowrap>List Unit Cost:</td>
		<td nowrap><input type="text" name="ListUnitCost" value="<%=((InitVendorInfo)?FormatCurrency(rsVendor.Fields.Item("fltList_Unit_Cost").Value):"")%>" size="10" readonly></td>
    </tr>
    <tr> 
		<td nowrap>Model Number:</td>
		<td nowrap colspan="3"><input type="text" name="ModelNumber" value="<%=((InitVendorInfo)?rsVendor.Fields.Item("chvModel_Number").Value:"")%>" readonly size="30" tabindex="5"></td>
    </tr>
</table>
<br><br>
<b style="font-size: 11pt; color: #8CAAE6;">Instance Information</b>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Inventory Name:</td>
		<td nowrap colspan="3"><input type="text" name="InventoryName" value="<%=Request.Form("InventoryName")%>" size="62" maxlength="80" tabindex="4"></td>
	</tr>
    <tr> 
		<td nowrap>Inventory Status:</td>
		<td nowrap><select name="Status" tabindex="5">
		<% 
		while (!rsStatus.EOF) { 			
		%>
			<option value="<%=(rsStatus.Fields.Item("insEquip_status_id").Value)%>" <%if (String(Request.Form("Status"))=="undefined") { Response.Write((rsStatus.Fields.Item("insEquip_status_id").Value=="1")?"SELECTED":"")} else { Response.Write((Request.Form("Status")==rsStatus.Fields.Item("insEquip_status_id").Value)?"SELECTED":"")}%>><%=(rsStatus.Fields.Item("chvStatusDesc").Value)%> 
		<% 
			rsStatus.MoveNext();
		} 
		%>
		</select></td>
		<td nowrap>Inventory ID:</td>
		<td nowrap><input type="text" name="InventoryID" value="<%=Request.Form("InventoryID")%>" size="10" maxlength="15" tabindex="6" onKeypress="AllowNumericOnly();" ></td>	  
    </tr>
    <tr> 
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="7">
			<option value="1" <%=((Request.Form("Type")=="1")?"SELECTED":"")%>>Requisition 
			<option value="2" <%=((Request.Form("Type")=="2")?"SELECTED":"")%>>Purchased Card Order 
			<option value="3" <%=((Request.Form("Type")=="3")?"SELECTED":"")%>>Donated 
        </select></td>		
		<td nowrap>Serial Number:</td>
		<td nowrap><input type="text" name="SerialNumber" value="<%=((String(Request.Form("SaveFields"))=="false")?Request.Form("SerialNumber"):"")%>" maxlength="20" size="10" tabindex="8"></td>	  		
    </tr>
    <tr> 
		<td nowrap>Date Ordered:</td>
		<td nowrap>
			<input type="text" name="DateOrdered" value="<%=Request.Form("DateOrdered")%>" size="11" maxlength="10" tabindex="9" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>PR Number:</td>
		<td nowrap><input type="text" name="PurchaseRequisitionNumber" value="<%=Request.Form("PurchaseRequisitionNumber")%>" maxlength="20" size="10" tabindex="11"></td>
    </tr>
	<tr>
		<td nowrap>Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Inventory Cost:</td>
		<td nowrap>$<input type="text" name="InventoryCost" value="<%if (InitVendorInfo) { Response.Write((rsVendor.Fields.Item("fltList_Unit_Cost").Value!="")?rsVendor.Fields.Item("fltList_Unit_Cost").Value:Request.Form("InventoryCost"))}%>" size="7" maxlength="7" tabindex="12" onKeypress="AllowNumericOnly();"></td>
	</tr>	
	<tr> 
		<td nowrap>Activation Key:</td>
		<td nowrap colspan="3"><input type="text" name="ActivationKey" value="<%=Request.Form("ActivationKey")%>" size="40" maxlength="50" tabindex="13" accesskey="L"></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="13" onClick="Save1(1);" class="btnstyle"></td>
		<td><input type="button" value="Save and Retain Fields" tabindex="14" onClick="Save1(2);" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="16" onClick="self.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="IsTemplate" value="0">
<input type="hidden" name="Initialized" value="true">
<input type="hidden" name="CInitialize" value="false">
<input type="hidden" name="MM_Insert" value="false">
<input type="hidden" name="SaveFields" value="false">
</form>
</body>
</html>
<%
rsAbstractClass.Close();
rsSubAbstractClass.Close();
rsStatus.Close();
if ((String(Request.Form("Initialized")) == "true") && (String(Request.Form("CInitialize")) == "false")) {
	rsConcreteClass.Close();
	rsVendor.Close();
}
%>
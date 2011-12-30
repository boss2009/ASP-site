<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update")) == "true"){
	var InventoryName = String(Request.Form("InventoryName")).replace(/'/g, "''");			
	var InventoryCost = ((String(Request.Form("InventoryCost"))=="")?"0":Request.Form("InventoryCost"));	
	var rsInventory = Server.CreateObject("ADODB.Recordset");
	rsInventory.ActiveConnection = MM_cnnASP02_STRING;
	rsInventory.Source = "{call dbo.cp_Update_EqCls_Inventory(" + Request.QueryString("intEquip_Set_id") + ",'" + Request.Form("SerialNumber") + "'," + Request.Form("PRNumber") + "," + InventoryCost +",'" + InventoryName + "'," + Request.Form("Status") + ",'" + Request.Form("OrderType") + "','" + Request.Form("DateOrdered") + "','" + Request.Form("DateReceived") + "','" + Request.Form("ActivationKey") + "',"+ Session("insStaff_id") + ",0)}";
	rsInventory.CursorType = 0;
	rsInventory.CursorLocation = 2;
	rsInventory.LockType = 3;
	rsInventory.Open();
	Response.Redirect("UpdateSuccessful2.asp?page=m003e0101.asp&intEquip_Set_id="+Request.QueryString("intEquip_Set_id"));
}

var rsInventory = Server.CreateObject("ADODB.Recordset");
rsInventory.ActiveConnection = MM_cnnASP02_STRING;
rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory(1,0,'',1," + Request.QueryString("intEquip_Set_id") + ",0)}";
rsInventory.CursorType = 0;
rsInventory.CursorLocation = 2;
rsInventory.LockType = 3;
rsInventory.Open();	

if ((rsInventory.Fields.Item("intRequisition_no").Value > 0) && (rsInventory.Fields.Item("intRequisition_no").Value < 30000)) {	 
	var rsPurchaseHeader = Server.CreateObject("ADODB.Recordset");
	rsPurchaseHeader.ActiveConnection = MM_cnnASP02_STRING;
	rsPurchaseHeader.Source = "{call dbo.cp_FrmHdr(14,"+rsInventory.Fields.Item("intRequisition_no").Value+")}";
	rsPurchaseHeader.CursorType = 0;
	rsPurchaseHeader.CursorLocation = 2;
	rsPurchaseHeader.LockType = 3;
	rsPurchaseHeader.Open();
	if (!rsPurchaseHeader.EOF) {
		Vendor = Trim(rsPurchaseHeader.Fields.Item("chvVendor").Value);
	} else {
		Vendor = "";
	}
} else {
	Vendor = "";
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
	<title>General Information</title>
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
				window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		if (Trim(document.frm0101.InventoryName.value)==""){
			alert("Enter Inventory Name.");
			document.frm0101.InventoryName.focus();
			return ;
		}
		if (isNaN(document.frm0101.InventoryCost.value)) {
			alert("Invalid Inventory Cost.");
			document.frm0101.InventoryCost.focus();
			return ;
		}
		if (!CheckDate(document.frm0101.DateOrdered.value)){
			alert("Invalid Ordered Date.");
			document.frm0101.DateOrdered.focus();
			return ;
		}
		if (!CheckDate(document.frm0101.DateReceived.value)){
			alert("Invalid Received Date");
			document.frm0101.DateReceived.focus();
			return ;
		}
		
		if (String(document.frm0101.CurrentInventoryStatus.value)!=String(document.frm0101.Status.value)) {
			switch (String(document.frm0101.Status.value)){
				//Loans
				case "2":
					alert("Cannot change inventory status.  Please assign this inventory to a Loan.");
					return ;
				break;
				case "4":
					alert("Cannot change inventory status.  Please assign this inventory to a Loan.");
					return ;
				break;
				case "25":
					alert("Cannot change inventory status.  Please assign this inventory to a Loan.");
					return ;
				break;
				case "26":
					alert("Cannot change inventory status.  Please assign this inventory to a Loan.");
					return ;
				break;
				case "19":
					alert("Cannot change inventory status.  Please assign this inventory to a Loan.");
					return ;
				break;
				case "20":
					alert("Cannot change inventory status.  Please assign this inventory to a Loan.");
					return ;
				break;
				case "21":
					alert("Cannot change inventory status.  Please assign this inventory to a Loan.");
					return ;
				break;
				case "3":
					alert("Cannot change inventory status.  Please assign this inventory to a Loan.");
					return ;
				break;

				//Buyouts
				case "11":
					alert("Cannot change inventory status.  Please assign this inventory to a Buyout.");					
					return ; 
				break;
				case "16":
					alert("Cannot change inventory status.  Please assign this inventory to a Buyout.");					
					return ; 
				break;
				case "14":
					alert("Cannot change inventory status.  Please assign this inventory to a Buyout.");					
					return ; 
				break;
				case "15":
					alert("Cannot change inventory status.  Please assign this inventory to a Buyout.");					
					return ; 
				break;
				case "17":
					alert("Cannot change inventory status.  Please assign this inventory to a Buyout.");					
					return ; 
				break;
				case "21":
					alert("Cannot change inventory status.  Please assign this inventory to a Buyout.");					
					return ; 
				break;
				//In Stock
				case "1":
					alert("Cannot change inventory status.  Please return the Loan/Buyout that the inventory is currently assigned to.");
					return ;
				break;
				//Under Repair
				case "6":
					alert("Cannot change inventory status.  Please assign this inventory to an Equipment Service.");					
					return ;					
				break;
				//Out of Inventory
				case "10":
					if (document.frm0101.CurrentInventoryStatus.value!="1") {
						alert("Cannot change inventory status to Out of Inventory unless it is In Stock.");
						return ;
					}
				break;
				case "8":
					if (document.frm0101.CurrentInventoryStatus.value!="1") {
						alert("Cannot change inventory status to Out of Inventory unless it is In Stock.");
						return ;
					}
				break;
				case "9":
					if (document.frm0101.CurrentInventoryStatus.value!="1") {
						alert("Cannot change inventory status to Out of Inventory unless it is In Stock.");
						return ;
					}
				break;
				case "7":
					if (document.frm0101.CurrentInventoryStatus.value!="1") {
						alert("Cannot change inventory status to Out of Inventory unless it is In Stock.");
						return ;
					}
				break;
				case "5":
					if (document.frm0101.CurrentInventoryStatus.value!="1") {
						alert("Cannot change inventory status to Out of Inventory unless it is In Stock.");
						return ;
					}
				break;
				default :
				break;
			}			
		}
		
		document.frm0101.submit();
	}
	</script>
</head>
<body onLoad="document.frm0101.InventoryName.focus();">
<form action="<%=MM_editAction%>" method="POST" name="frm0101">
<b style="font-size: 11pt; color: #8CAAE6;">Class Information</b>
<table cellpadding="1" cellspacing="1">
	<tr>     	
		<td nowrap>Abstract Class:</td>		
		<td nowrap width="200"><input type="text" name="AbstractClass" maxlength="30" value="<%=rsInventory.Fields.Item("chvEqCls_name_2").Value%>" size="30" readonly tabindex="1" accesskey="F"></td>
		<td nowrap>Subject To:</td>
		<td nowrap><select name="SubjectTo" tabindex="6" disabled>
			<option value="0" <%=((rsInventory.Fields.Item("chvSbjTotax").Value == "0")?"SELECTED":"")%>>No Tax 
			<option value="1" <%=((rsInventory.Fields.Item("chvSbjTotax").Value == "1")?"SELECTED":"")%>>PST 
			<option value="2" <%=((rsInventory.Fields.Item("chvSbjTotax").Value == "2")?"SELECTED":"")%>>GST 
			<option value="3" <%=((rsInventory.Fields.Item("chvSbjTotax").Value == "3")?"SELECTED":"")%>>Both 
        </select></td>	  
	</tr>
	<tr> 
		<td nowrap>Sub Abstract Class:</td>
		<td nowrap><input type="text" name="SubAbstractClass" maxlength="30" value="<%=rsInventory.Fields.Item("chvEqCls_name_1").Value%>" size="30" readonly tabindex="2"></td>
		<td nowrap>Parts Warranty:</td>
		<td nowrap><select name="PartsWarrantyLength" tabindex="7" disabled>
		<% 
		while (!rsWarrantyLength.EOF) { 
		%>
			<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" <%=((rsInventory.Fields.Item("insPartsWLen").Value==rsWarrantyLength.Fields.Item("insWarrenty_id").Value)?"SELECTED":"")%>><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%> 
		<% 
			rsWarrantyLength.MoveNext();			
		} 
		%>
		</select></td>	  
	</tr>
    <tr> 
		<td nowrap>Concrete Class:</td>
		<td nowrap><input type="text" name="ConcreteClass" value="<%=rsInventory.Fields.Item("chvEqCls_name").Value%>" size="30" maxlength="30" readonly tabindex="3"></td>
		<td nowrap>Labour Warranty:</td>
		<td nowrap><select name="LabourWarrantyLength" tabindex="8" disabled>
			<% 
			rsWarrantyLength.MoveFirst();			
			while (!rsWarrantyLength.EOF) { 			
			%>
				<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" <%=((rsInventory.Fields.Item("insLaborWLen").Value==rsWarrantyLength.Fields.Item("insWarrenty_id").Value)?"SELECTED":"")%>><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%> 
			<% 
				rsWarrantyLength.MoveNext();
			} 
			%>
        </select></td>	  
    </tr>
	<tr>		
		<td nowrap>Default Vendor:</td>		
		<td nowrap><input type="text" name="DefaultVendor" value="<%=rsInventory.Fields.Item("chvVendor_Name").Value%>" size="30" readonly tabindex="4" maxlength="30"></td>
		<td nowrap>List Unit Cost:</td>		
		<td nowrap><input type="text" name="ListUnitCost" value="<%=FormatCurrency(rsInventory.Fields.Item("fltList_Unit_Cost").Value)%>" size="10" readonly tabindex="9"></td>
	</tr>
    <tr> 
		<td nowrap>Model Number:</td>
		<td nowrap><input type="text" name="ModelNumber" value="<%=(rsInventory.Fields.Item("chvModel_Number").Value)%>" maxlength="50" readonly size="30" tabindex="5"></td>
		<td nowrap colspan="2"></td>
    </tr>
</table>
<br>
<b style="font-size: 11pt; color: #8CAAE6;">Instance Information</b>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Inventory Name:</td>
		<td nowrap colspan="3"><input type="text" value="<%=rsInventory.Fields.Item("chvInventory_Name").Value%>" maxlength="62" size="62" tabindex="10" name="InventoryName">
	</td>	  
	</tr>	
    <tr> 
		<td nowrap>Inventory Status:</td>
		<td nowrap><select name="Status" tabindex="11">
		<% 
		rsStatus.MoveFirst();			
		while (!rsStatus.EOF) { 			
		%>
			<option value="<%=(rsStatus.Fields.Item("insEquip_status_id").Value)%>" <%=((rsInventory.Fields.Item("insCurrent_Status").Value==rsStatus.Fields.Item("insEquip_status_id").Value)?"SELECTED":"")%>><%=(rsStatus.Fields.Item("chvStatusDesc").Value)%> 
		<% 
			rsStatus.MoveNext();
		} 
		%>
        </select></td>
		<td nowrap>Inventory ID:</td>
		<td nowrap><input type="text" name="InventoryID" value="<%=rsInventory.Fields.Item("intBar_Code_no").Value%>" size="10" maxlength="15" tabindex="12" style="border: none" readonly></td>	  
    </tr>
    <tr> 
		<td nowrap>Order Type:</td>
		<td nowrap><select name="OrderType" tabindex="13">
			<option value="1" <%=((rsInventory.Fields.Item("chrObtain_by").Value=="1")?"SELECTED":"")%>>Requisition 
			<option value="2" <%=((rsInventory.Fields.Item("chrObtain_by").Value=="2")?"SELECTED":"")%>>Purchased Card Order 
			<option value="3" <%=((rsInventory.Fields.Item("chrObtain_by").Value=="3")?"SELECTED":"")%>>Donated 
		</select></td>
		<td nowrap>Serial Number:</td>
		<td nowrap><input type="text" name="SerialNumber" value="<%=(rsInventory.Fields.Item("chvSerial_Number").Value)%>" maxlength="40" size="20" tabindex="14"></td>	  
    </tr>
    <tr> 
		<td nowrap>Date Ordered:</td>
		<td nowrap>
			<input type="text" name="DateOrdered" value="<%=FilterDate(rsInventory.Fields.Item("dtsOrd_Date").Value)%>" size="11" maxlength=10 tabindex="15" readonly style="border: none" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>PR Number:</td>
		<td nowrap><input type="text" name="PRNumber" value="<%=(rsInventory.Fields.Item("intRequisition_no").Value)%>" maxlength="20" size="10" tabindex="16"></td>
    </tr>
    <tr> 
		<td nowrap>Date Received:</td>
		<td nowrap>
		<%
		var UseInventoryDate = true;
		if (rsInventory.Fields.Item("intRequisition_no").Value > 0) {
			if (!((rsInventory.Fields.Item("intRequisition_no").Value == 9999) || (rsInventory.Fields.Item("intRequisition_no").Value == 99999) || (rsInventory.Fields.Item("intRequisition_no").Value == 999999))) {
				UseInventoryDate = false;
			}
		}
		%>

<%
// + Nov.08.2005
   	if (!rsInventory.EOF) {
       Response.Write("<input type='text' name='DateReceived' value='");
       if (UseInventoryDate==true) {
			Response.Write("FilterDate(rsInventory.Fields.Item('dtsIvtry_Rec_Date').Value)");
	   } else {
			Response.Write("FilterDate(rsInventory.Fields.Item('dtsRec_Date').Value)");
	   }
       Response.Write("' size='11' maxlength=10 tabindex='17' readonly style='border: none' onChange='FormatDate(this)' ");
	}
%>

    	    <span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Vendor:</td>
		<td nowrap><input type="text" name="Vendor" readonly value="<%=Vendor%>" style="border: none" tabindex="18" size="40"></td>
    </tr>
    <tr>
		<td nowrap>Activation Key:</td>
		<td nowrap><input type="text" name="ActivationKey" value="<%=rsInventory.Fields.Item("chvActivekey").Value%>" size="40" maxlength="50" tabindex="19"></td>
		<td nowrap>Inventory Cost:</td>
		<td nowrap>$<input type="text" name="InventoryCost" value="<%=((rsInventory.Fields.Item("fltPurchase_Cost").Value=="")?"0":rsInventory.Fields.Item("fltPurchase_Cost").Value)%>" size="10" maxlength="7" tabindex="20"></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="21" class="btnstyle"></td>
 		<td><input type="reset" value="Undo Changes" tabindex="22" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="23" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="intEquip_Set_id" value="<%=Request.QueryString("intEquip_Set_id")%>">
<input type="hidden" name="ActionDelete" value="false">
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="CurrentInventoryStatus" value="<%=rsInventory.Fields.Item("insCurrent_Status").Value%>">
</form>
</body>
</html>
<%
rsInventory.Close();
rsWarrantyLength.Close()
rsStatus.Close();
%>
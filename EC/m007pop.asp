<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_edit")) == "true"){
	var rsVendorDeal = Server.CreateObject("ADODB.Recordset");
	rsVendorDeal.ActiveConnection = MM_cnnASP02_STRING;
	rsVendorDeal.Source = "{call dbo.cp_Update_EqCls_Dtl("+Request.Form("ClassID")+","+Request.Form("VendorID")+",'"+Request.Form("EntryDate")+"',"+Request.Form("ListUnitCost")+","+Request.Form("PriceQuantity")+","+((Request.Form("IsDefaultSupplier")=="on")?"1":"0")+",'"+ Request.Form("ContractPONumber")+"',"+Request.Form("PartsWarrantyLength")+","+Request.Form("LabourWarrantyLength")+","+Request.Form("intEqCls_Dtl_id")+",0)}";
	rsVendorDeal.CursorType = 0;
	rsVendorDeal.CursorLocation = 2;
	rsVendorDeal.LockType = 3;
	rsVendorDeal.Open();
	Response.Redirect("AddDeleteSuccessful.asp");
}

var rsVendorDeal = Server.CreateObject("ADODB.Recordset");
rsVendorDeal.ActiveConnection = MM_cnnASP02_STRING;
rsVendorDeal.Source = "{call dbo.cp_Get_EqCls_Vendor(0," + Request.QueryString("intEqCls_Dtl_id") + ",1)}";	
rsVendorDeal.CursorType = 0;
rsVendorDeal.CursorLocation = 2;
rsVendorDeal.LockType = 3;
rsVendorDeal.Open();

var rsWarrantyLength = Server.CreateObject("ADODB.Recordset");
rsWarrantyLength.ActiveConnection = MM_cnnASP02_STRING;
rsWarrantyLength.Source = "{call dbo.cp_ASP_lkup(62)}";
rsWarrantyLength.CursorType = 0;
rsWarrantyLength.CursorLocation = 2;
rsWarrantyLength.LockType = 3;
rsWarrantyLength.Open();
	
var rsPriceQuantity = Server.CreateObject("ADODB.Recordset");
rsPriceQuantity.ActiveConnection = MM_cnnASP02_STRING;
rsPriceQuantity.Source = "{call dbo.cp_ASP_lkup(73)}";
rsPriceQuantity.CursorType = 0;
rsPriceQuantity.CursorLocation = 2;
rsPriceQuantity.LockType = 3;
rsPriceQuantity.Open();		
%>
<html>
<head>
	<title>Vendor Deal</title>
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
		document.frm0103d.submit();
	}
	</script>
</head>
<body onLoad="document.frm0103d.IsDefaultSupplier.focus();">
<form action="<%=MM_editAction%>" method="POST" name="frm0103d">
<h5>Vendor Deal</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Default Supplier:</td>
		<td nowrap><input type="checkbox" name="IsDefaultSupplier" <%=((rsVendorDeal.Fields.Item("bitIsCurrent").Value)?"CHECKED":"")%> accesskey="F" tabindex="1" class="chkstyle"></td>
	</tr>
	<tr> 
		<td nowrap>Company Name:</td>
		<td nowrap><input type="text" name="CompanyName" value="<%=(rsVendorDeal.Fields.Item("chvCompany_Name").Value)%>" readonly tabindex="2" size="40"></td>
	</tr>
	<tr> 
		<td nowrap valign="top">Street Address:</td>
		<td nowrap valign="top"><textarea name="StreetAddress" readonly tabindex="3"><%=(rsVendorDeal.Fields.Item("chvAddress").Value)%></textarea></td>
	</tr>
	<tr> 
		<td nowrap>City:</td>
		<td nowrap><input type="text" name="City" value="<%=(rsVendorDeal.Fields.Item("chvCity").Value)%>" readonly tabindex="4"></td>
	</tr>
	<tr> 
		<td nowrap>Province/State:</td>
		<td nowrap><input type="text" name="ProvinceState" value="<%=(rsVendorDeal.Fields.Item("chrprvst_abbv").Value)%>" readonly tabindex="5" size="2"></td>
	</tr>
	<tr> 
		<td nowrap>Phone Number:</td>		
		<td nowrap><%=FormatPhoneNumber(rsVendorDeal.Fields.Item("chvPhone_Type").Value,rsVendorDeal.Fields.Item("chvPhone1_Arcd").Value,rsVendorDeal.Fields.Item("chvPhone1_Num").Value,rsVendorDeal.Fields.Item("chvPhone1_Ext").Value,"","","","","","","","")%></td>
	</tr>
	<tr> 
		<td nowrap>Entry Date:</td>
		<td nowrap>
			<input type="text" name="EntryDate" value="<%=FilterDate(rsVendorDeal.Fields.Item("dtsEntry_Date").Value)%>" size="11" maxlength=10 readonly tabindex="6" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td nowrap>Contract PO:</td>
		<td nowrap><input type="text" name="ContractPONumber" value="<%=(rsVendorDeal.Fields.Item("chvContract_PO").Value)%>" tabindex="7"></td>
	</tr>	
	<tr> 
		<td nowrap>List Unit Cost:</td>
		<td nowrap>$<input type="text" name="ListUnitCost" value="<%=(rsVendorDeal.Fields.Item("fltList_Unit_Cost").Value)%>" tabindex="8" size="6"></td>
	</tr>
	<tr> 
		<td nowrap>Price Quantity:</td>
		<td nowrap><select name="PriceQuantity" tabindex="9">
			<% 
			while (!rsPriceQuantity.EOF) { 
			%>
				<option value="<%=(rsPriceQuantity.Fields.Item("insPrice_qty_id").Value)%>" <%=((rsPriceQuantity.Fields.Item("insPrice_qty_id").Value==rsVendorDeal.Fields.Item("insPrice_Qty_Id").Value)?"SELECTED":"")%>><%=(rsPriceQuantity.Fields.Item("chvName").Value)%>
			<% 
			rsPriceQuantity.MoveNext();			
			} 
			%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap>Parts Warranty Length:</td>
		<td nowrap><select name="PartsWarrantyLength" tabindex="10">
			<% 
			while (!rsWarrantyLength.EOF) { 
			%>
				<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" <%=((rsVendorDeal.Fields.Item("insPartsWLen").Value==rsWarrantyLength.Fields.Item("insWarrenty_id").Value)?"SELECTED":"")%>><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%>
			<% 
			rsWarrantyLength.MoveNext();			
			} 
			%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap>Labour Warranty Length:</td>
		<td nowrap><select name="LabourWarrantyLength" tabindex="11" accesskey="L">	  
			<% 
			rsWarrantyLength.MoveFirst();			
			while (!rsWarrantyLength.EOF) { 			
			%>
				<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" <%=((rsVendorDeal.Fields.Item("insLaborWLen").Value==rsWarrantyLength.Fields.Item("insWarrenty_id").Value)?"SELECTED":"")%>><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%>
			<% 
			rsWarrantyLength.MoveNext();
			} 
			%>
		</select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="12" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close()" tabindex="13" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="VendorID" value="<%=rsVendorDeal.Fields.Item("insVendor_id").Value%>">
<input type="hidden" name="ClassID" value="<%=Request.QueryString("ClassID")%>">
<input type="hidden" name="intEqCls_Dtl_id" value="<%=Request.QueryString("intEqCls_Dtl_id")%>">
<input type="hidden" name="MM_edit" value="true">
</form>
</body>
</html>
<%
rsVendorDeal.Close();
rsWarrantyLength.Close();
rsPriceQuantity.Close();
%>
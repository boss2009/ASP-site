<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_insertAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_insertAction += "?" + Request.QueryString;
}

if (Request.Form("MM_insert") == "true"){
	var CompanyName = String(Request.Form("CompanyName")).replace(/'/g, "''");		
	var rsDeal = Server.CreateObject("ADODB.Recordset");
	rsDeal.ActiveConnection = MM_cnnASP02_STRING;
	rsDeal.Source = "{call dbo.cp_Insert_EqpCls_Dtl("+Request.Form("ClassID")+","+CompanyName+",'"+Request.Form("EntryDate")+"',"+Request.Form("ListUnitCost")+","+Request.Form("PriceQuantity")+","+((Request.Form("IsDefaultSupplier")=="on")?"1":"0")+",'"+ Request.Form("ContractPONumber")+"',"+Request.Form("PartsWarrantyLength")+","+Request.Form("LabourWarrantyLength")+",0)}";
	rsDeal.CursorType = 0;
	rsDeal.CursorLocation = 2;
	rsDeal.LockType = 3;
	rsDeal.Open();
	Response.Redirect("InsertSuccessful.html");
}

var rsCompany = Server.CreateObject("ADODB.Recordset");
rsCompany.ActiveConnection = MM_cnnASP02_STRING;
rsCompany.Source = "{call dbo.cp_company2(0,'',0,0,0,0,0,1,0,'',0,'Q',0)}";
rsCompany.CursorType = 0;
rsCompany.CursorLocation = 2;
rsCompany.LockType = 3;
rsCompany.Open();

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
	<title>New Vendor Deal</title>
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
<form action="<%=MM_insertAction%>" method="POST" name="frm0103d">
<h5>New Vendor Deal</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Default Supplier:</td>
<!--	<td nowrap><input type="checkbox" name="IsDefaultSupplier" accesskey="F" tabindex="1" <%=((Session("MM_UserAuthorization") < 5)?"disabled":"")%> class="chkstyle"></td>-->
		<td nowrap><input type="checkbox" name="IsDefaultSupplier" accesskey="F" tabindex="1" class="chkstyle"></td>
	</tr>
	<tr> 
		<td nowrap>Company Name:</td>
		<td nowrap><select name="CompanyName" tabindex="2">
		<%
		while (!rsCompany.EOF){
		%>
			<option value="<%=(rsCompany.Fields.Item("intCompany_id").Value)%>"><%=(rsCompany.Fields.Item("chvCompany_Name").Value)%>
		<%
			rsCompany.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap>Entry Date:</td>
		<td nowrap>
			<input type="text" name="EntryDate" value="<%=CurrentDate()%>" size="11" maxlength="10" onKeypress="AllowNumericOnly();" tabindex="3" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td nowrap>Contract PO:</td>
		<td nowrap><input type="text" name="ContractPONumber" tabindex="4"></td>
	</tr>	
	<tr> 
		<td nowrap>List Unit Cost:</td>
		<td nowrap>$<input type="text" name="ListUnitCost" value="0.00" tabindex="5" size="6" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr> 
		<td nowrap>Price Quantity:</td>
		<td nowrap><select name="PriceQuantity" tabindex="6">
			<% 
			while (!rsPriceQuantity.EOF) { 
			%>
				<option value="<%=(rsPriceQuantity.Fields.Item("insPrice_qty_id").Value)%>" ><%=(rsPriceQuantity.Fields.Item("chvName").Value)%>
			<% 
			rsPriceQuantity.MoveNext();			
			} 
			%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap>Parts Warranty Length:</td>
		<td nowrap><select name="PartsWarrantyLength" tabindex="7">
			<% 
			while (!rsWarrantyLength.EOF) { 
			%>
				<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" ><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%>
			<% 
			rsWarrantyLength.MoveNext();			
			} 
			%>
		</select></td>
	</tr>
	<tr> 
		<td nowrap>Labour Warranty Length:</td>
		<td nowrap><select name="LabourWarrantyLength" tabindex="8" accesskey="L">	  
			<% 
			rsWarrantyLength.MoveFirst();			
			while (!rsWarrantyLength.EOF) { 			
			%>
				<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>"><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%>
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
		<td><input type="button" value="Save" onClick="Save();" tabindex="9" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="10" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ClassID" value="<%=Request.QueryString("ClassID")%>">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsCompany.Close();
rsWarrantyLength.Close();
rsPriceQuantity.Close();	
%>
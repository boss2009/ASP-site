<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsSysOptr = Server.CreateObject("ADODB.Recordset");
rsSysOptr.ActiveConnection = MM_cnnASP02_STRING;
rsSysOptr.Source = "{call dbo.cp_SysOptr(0,0,22)}";
rsSysOptr.CursorType = 0;
rsSysOptr.CursorLocation = 2;
rsSysOptr.LockType = 3;
rsSysOptr.Open();

var rsVendor = Server.CreateObject("ADODB.Recordset");
rsVendor.ActiveConnection = MM_cnnASP02_STRING;
rsVendor.Source = "{call dbo.cp_company2(0,'',0,0,0,0,0,1,0,'',0,'Q',0)}";
rsVendor.CursorType = 0;
rsVendor.CursorLocation = 2;
rsVendor.LockType = 3;
rsVendor.Open();
%>
<html>
<head>
	<title>Equipment Class - Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m007Srh01.js"></script>
	<script language="javascript">
	if (window.focus) self.focus();
	function initscr() {  
		ockbTax.style.visibility="hidden";
		oOptr3.style.visibility="hidden";  
		document.frm07s01.Operator.focus();
	}

	function SelOpt() {	
		var len = document.frm07s01.SearchType.length;
		var Idparam = 1;                 // init.
	
		for (var i=0;i <len; i++){
			if (document.frm07s01.SearchType[i].checked) Idparam = i;
		}
		switch ( Idparam ) {
			// Name
			case 0: 
				oOptr1.style.visibility="visible";
				ockbTax.style.visibility="hidden";
				oOptr3.style.visibility="hidden";
			break;
			//Tax
			case 1: 
				oOptr1.style.visibility="hidden";
				ockbTax.style.visibility="visible";
				oOptr3.style.visibility="hidden";
			break;
			//Vendors
			case 2: 
				oOptr1.style.visibility="hidden";
				ockbTax.style.visibility="hidden";	   		   
				oOptr3.style.visibility="visible";
			break; 
			default: 
				oOptr1.style.visibility="visible";
				ockbTax.style.visibility="hidden";	   		   
				oOptr3.style.visibility="hidden";
			break;
		}       
	}

	function CnstrFltr(output) {	
		var len = document.frm07s01.SearchType.length;
		var stgPgQuery = "";
		var Idparam = 1;	
		for (var i=0;i <len; i++){
			if (document.frm07s01.SearchType[i].checked) Idparam = i;
		}

		switch (Idparam) {
			//Name
			case 0: 
				var chvOprd = "1" ; 
				var chrNot  = "";
				var chvOptr = document.frm07s01.Operator[document.frm07s01.Operator.selectedIndex].value ;
				var chvStg1 = document.frm07s01.ClassName.value;
			break;
			//Tax
			case 1: 
				var chvOprd = "2" ; 
				var chrNot  = ""   
				var chvOptr = "" ;
				var chvStg1= "";
				var intTaxlen = document.frm07s01.TaxType.length;
				for (var j=0;j <intTaxlen; j++){
					if (document.frm07s01.TaxType[j].selected) chvStg1 = document.frm07s01.TaxType[j].value;
				}
			break;
			//Vendor
			case 2: 
				var chvOprd = "3" ; 
				var chrNot  = "";
				var chvOptr = "" ;
				var chvStg1 = document.frm07s01.VendorName[document.frm07s01.VendorName.selectedIndex].value ;
			break; 
			default: 
				var chvOprd = "1" ; 
				var chrNot  = "";
				var chvOptr = document.frm07s01.Operator[document.frm07s01.Operator.selectedIndex].value ;
				var chvStg1 = document.frm07s01.ClassName.value;
			break;			
		}       
		var chvStg2 = "";
		var stgFilter = ACfltr_07(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
		var inspSrtBy  = document.frm07s01.SortBy.value;
		var inspSrtOrd  = document.frm07s01.OrderBy.value;	
		if (output ==1) {
			document.frm07s01.action = "m007q02.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;
		} else {
			document.frm07s01.action = "m007q02excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter;
			document.frm07s01.target = "_blank"; 			
			document.frm07s01.submit() ; 			
		}
	}
	</script>
</head>
<body onload="initscr();">
<form name="frm07s01" method="post" onSubmit="CnstrFltr(1);">
<h3>Equipment Class - Search</h3>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap width="100"><input type="radio" name="SearchType" value="1" checked accesskey="F" tabindex="1" onClick="SelOpt()" class="chkstyle">Class Name</td>
		<td nowrap><DIV ID="oOptr1" STYLE="visibility:visible"> 
			<select name="Operator" tabindex="2">
			<% 
			while (!rsSysOptr.EOF) { 
			%>
				<option value="<%=(rsSysOptr.Fields.Item("intOptrId").Value)%>" <%=((rsSysOptr.Fields.Item("intOptrId").Value == 1)?"SELECTED":"")%> ><%=(rsSysOptr.Fields.Item("chvOptrDesc").Value)%></option>
            <% 	
				rsSysOptr.MoveNext(); 
			}
            %>
			</select>
			<input type="text" name="ClassName" size="20" tabindex="3" maxlength="20">
		</DIV></td>
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap width="100"><input type="radio" name="SearchType" value="2" tabindex="4" onClick="SelOpt()" class="chkstyle">Subject To</td>
		<td nowrap><DIV ID="ockbTax" STYLE="visibility:visible">
			<select name="TaxType" tabindex="5">
				<option value="0">No Tax
				<option value="1">GST
				<option value="2">PST
				<option value="3">Both									
			</select></div>
		</td>
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap width="100"><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="6" class="chkstyle">Vendor Name</td>
		<td nowrap><div id="oOptr3" style="visibility:visible"> 
			<select name="VendorName" tabindex="7">
            <% 
			while (!rsVendor.EOF) { 
			%>
            	<option value="<%=(rsVendor.Fields.Item("intCompany_id").Value)%>" <%=((rsVendor.Fields.Item("intCompany_id").Value == 1)?"SELECTED":"")%>><%=(rsVendor.Fields.Item("chvCompany_Name").Value)%>
            <% 
				rsVendor.MoveNext();   
			}
			%>
			</select>
		</div></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td colspan="3">
			Sort By
			<select name="SortBy" tabindex="8">
				<option value="1">Class Name</option>		
				<option value="0">Class ID</option>
			</select>
			Order
			<select name="OrderBy" tabindex="9" accesskey="L">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="submit" value="Search" tabindex="10" class="btnstyle"></td>
		<td><input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="11" class="btnstyle"></td>
	</tr>
</table>
</form>
</body>
</html>
<%
rsSysOptr.Close();
rsVendor.Close();
%>
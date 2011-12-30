<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.QueryString("Search"))=="true") {
	var rsEquipmentClass__inspSrtBy = "1";
	if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
		rsEquipmentClass__inspSrtBy = String(Request.QueryString("inspSrtBy"));
	}
	var rsEquipmentClass__inspSrtOrd = "0";
	if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
		rsEquipmentClass__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
	}
	var rsEquipmentClass__chvFilter = "";
	if(String(Request.QueryString("chvFilter")) != "undefined") { 
		rsEquipmentClass__chvFilter = String(Request.QueryString("chvFilter"));
	}
	var rsEquipmentClass = Server.CreateObject("ADODB.Recordset");
	rsEquipmentClass.ActiveConnection = MM_cnnASP02_STRING;
	rsEquipmentClass.Source = "{call dbo.cp_EC_Eqp_Class("+ rsEquipmentClass__inspSrtBy.replace(/'/g, "''") + ","+ rsEquipmentClass__inspSrtOrd.replace(/'/g, "''") + ",'"+ rsEquipmentClass__chvFilter.replace(/'/g, "''") + "')}";
	rsEquipmentClass.CursorType = 0;
	rsEquipmentClass.CursorLocation = 2;
	rsEquipmentClass.LockType = 3;
	rsEquipmentClass.Open();
}

var rsSysOptr = Server.CreateObject("ADODB.Recordset");
rsSysOptr.ActiveConnection = MM_cnnASP02_STRING;
rsSysOptr.Source = "{call dbo.cp_SysOptr(0,0,22)}";
rsSysOptr.CursorType = 0;
rsSysOptr.CursorLocation = 2;
rsSysOptr.LockType = 3;
rsSysOptr.Open();
%>
<html>
<head>
	<title>Inventory Class Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m007Srh01.js"></script>
	<script language="javascript">
	if (window.focus) self.focus();
	function CnstrFltr() {
		var chvOptr = document.frm0101.ClassNameOperator[document.frm0101.ClassNameOperator.selectedIndex].value ;
		var chvStg1 = document.frm0101.ClassName.value;
		var stgFilter = ACfltr_07("1","",chvOptr,chvStg1,"");
		stgPgQuery = "m008p0401.asp?";
		stgPgQuery   += "Search=true&inspSrtBy=1&inspSrtOrd=0&chvFilter=" + stgFilter ;
		document.frm0101.action = stgPgQuery ;
		document.frm0101.submit() ; 
	}
	
	function SelectClass(){
		if (document.frm0101.SearchResult.selectedIndex==-1){
			alert("Select a class.")
			return ;
		}
		if (!top.opener.closed) {	
			top.opener.document.frm08s01.ClassSearchText.value=document.frm0101.SearchResult.options[document.frm0101.SearchResult.selectedIndex].text;
			top.opener.document.frm08s01.ClassSearchID.value=document.frm0101.SearchResult[document.frm0101.SearchResult.selectedIndex].value;
		}			
		top.window.close();
	}
	
	function init(){
	<% 
	if (String(Request.QueryString("Search")) == "true") { 
	%>
		document.frm0101.SearchResult.focus();
	<% 
	} else { 
	%>
		document.frm0101.ClassNameOperator.focus();
	<% 
	} 
	%>
	}	
	</script>
</head>
<body onLoad="init();">
<form name="frm0101" method="POST" action="">
<h5>Search Criteria</h5>
<table cellpadding="1" cellspacing="1">	
	<tr>
		<td nowrap>
			Class Name:
			<select name="ClassNameOperator" tabindex="1" accesskey="F">
			<% 
			while (!rsSysOptr.EOF) { 
			%>
				<option value="<%=(rsSysOptr.Fields.Item("intOptrId").Value)%>" <%=((rsSysOptr.Fields.Item("intOptrId").Value == 2)?"SELECTED":"")%> ><%=(rsSysOptr.Fields.Item("chvOptrDesc").Value)%></option>
			<% 	
				rsSysOptr.MoveNext(); 
			}
			%>
			</select>
			<input type="text" name="ClassName" size="40" tabindex="2" maxlength="50" value="<%=Request.QueryString("ClassName")%>">
			<input type="button" value="Search" onClick="CnstrFltr();" tabindex="3" class="btnstyle">
		</td>		
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><select name="SearchResult" size="20" style="width:420px" tabindex="4">
<% 
	if (String(Request.QueryString("Search")) == "true") { 
		var count = 0;
		while (!rsEquipmentClass.EOF) {
			count++;
			switch(rsEquipmentClass.Fields.Item("chrClass_Type").Value){ 
				case 'C':
%>
				<option value="<%=(rsEquipmentClass.Fields.Item("insEquip_Class_id").Value)%>"><%=(rsEquipmentClass.Fields.Item("chvClass_Name").Value)%> - Concrete
<%
				break;
			}
			rsEquipmentClass.MoveNext();
		}
	}
%>
		</select></td>		
	</tr>
	<tr>
		<td nowrap><input type="button" value="Select Class" tabindex="5" onClick="SelectClass();" class="btnstyle"></td>
	</tr>
</table>
<% 
if (String(Request.QueryString("Search")) == "true") { 
	if (count > 0) rsEquipmentClass.MoveFirst();
		while (!rsEquipmentClass.EOF) {
%>
<input type="hidden" name="ListUnitCost" value="<%=(rsEquipmentClass.Fields.Item("fltList_Unit_Cost").Value)%>">
<%
			rsEquipmentClass.MoveNext();
		}
	rsEquipmentClass.Close();
}
%>
</form>
</body>
</html>
<%
rsSysOptr.Close();
%>
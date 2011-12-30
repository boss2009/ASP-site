<!--------------------------------------------------------------------------
* File Name: m012p0102.asp
* Title: Inventory Class Search
* Main SP: cp_EC_Eqp_Class
* Description: This page is used to search for an inventory class.
* Author: T.H
--------------------------------------------------------------------------->
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
		var chvOptr = document.frm12p02.ClassNameOperator[document.frm12p02.ClassNameOperator.selectedIndex].value ;
		var chvStg1 = document.frm12p02.ClassName.value;
		var stgFilter = ACfltr_07("1","",chvOptr,chvStg1,"");
		stgPgQuery = "m012p0102.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&intEqpRequest_id=<%=Request.QueryString("intEqpRequest_id")%>&";
		stgPgQuery   += "Search=true&inspSrtBy=1&inspSrtOrd=0&chvFilter=" + stgFilter ;
		document.frm12p02.action = stgPgQuery ;
		document.frm12p02.submit() ; 
	}
	
	function SelectClass(){
		if (document.frm12p02.SearchResult.selectedIndex==-1){
			alert("Select a class.")
			return ;
		}			
		if (!top.opener.closed) {		
	<% 
	if (String(Request.QueryString("intEqpRequest_id"))=="") { 
	%>
			top.opener.document.frm0203.action = "m012a0203.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&intEqpRequest_id=<%=Request.QueryString("intEqpRequest_id")%>&ClassID="+document.frm12p02.SearchResult[document.frm12p02.SearchResult.selectedIndex].value+"&ClassName="+document.frm12p02.SearchResult.options[document.frm12p02.SearchResult.selectedIndex].text;	
	<% 
	} else { 
	%>
			top.opener.document.frm0203.action = "m012e0201.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&intEqpRequest_id=<%=Request.QueryString("intEqpRequest_id")%>&ClassID="+document.frm12p02.SearchResult[document.frm12p02.SearchResult.selectedIndex].value+"&ClassName="+document.frm12p02.SearchResult.options[document.frm12p02.SearchResult.selectedIndex].text;
	<% 
	} 
	%>
			top.opener.document.frm0203.submit();
		}
		top.window.close();
	}	

	function init(){
	<% 
	if (String(Request.QueryString("Search")) == "true") { 
	%>
		document.frm12p02.SearchResult.focus();
	<% 
	} else { 
	%>
		document.frm12p02.ClassNameOperator.focus();
	<% 
	} 
	%>
	}	
	</script>
</head>
<body onLoad="init();">
<form name="frm12p02" method="POST">
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
		<td><select name="SearchResult" size="20" style="width:420px" tabindex="4">
<% 
	var count = 0;
	if (String(Request.QueryString("Search")) == "true") { 
		while (!rsEquipmentClass.EOF) {
			switch(rsEquipmentClass.Fields.Item("chrClass_Type").Value){ 
				case 'A': 
%>
				<option value="<%=(rsEquipmentClass.Fields.Item("insEquip_Class_id").Value)%>"><%=FilterQuotes(rsEquipmentClass.Fields.Item("chvClass_Name").Value)%> - Abstract
<%
				break;
				case 'S':
%>
				<option value="<%=(rsEquipmentClass.Fields.Item("insEquip_Class_id").Value)%>"><%=FilterQuotes(rsEquipmentClass.Fields.Item("chvClass_Name").Value)%> - Sub Abstract
<%
				break;
				case 'C':
%>
				<option value="<%=(rsEquipmentClass.Fields.Item("insEquip_Class_id").Value)%>"><%=FilterQuotes(rsEquipmentClass.Fields.Item("chvClass_Name").Value)%> - Concrete
<%
				break;
			}
			count++;
			rsEquipmentClass.MoveNext();
		}
	}
%>
		</select></td>
	</tr>
	<tr>
		<td>
			<input type="button" value="Select Class" tabindex="5" onClick="SelectClass();" <%=((count==0)?"DISABLED":"")%> class="btnstyle">
			<input type="button" value="Close" tabindex="6" onClick="self.close();" class="btnstyle">
		</td>
	</tr>
</table>
</form>
</body>
</html>
<%
rsSysOptr.Close();
%>
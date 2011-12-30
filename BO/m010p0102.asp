<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.QueryString("Search"))=="true") {
	var rsBundle__inspSrtBy = "1";
	if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
	  rsBundle__inspSrtBy = String(Request.QueryString("inspSrtBy"));
	}
	var rsBundle__inspSrtOrd = "0";
	if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
	  rsBundle__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
	}
	var rsBundle__chvFilter = "";
	if(String(Request.QueryString("chvFilter")) != "undefined") { 
	  rsBundle__chvFilter = String(Request.QueryString("chvFilter"));
	}
	var rsBundle = Server.CreateObject("ADODB.Recordset");
	rsBundle.ActiveConnection = MM_cnnASP02_STRING;
	rsBundle.Source = "{call dbo.cp_Bundle(0,'',0.0,0,1,1,'',0,"+Session("insStaff_id")+","+rsBundle__inspSrtBy.replace(/'/g, "''")+","+rsBundle__inspSrtOrd.replace(/'/g, "''")+",'"+rsBundle__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
	rsBundle.CursorType = 0;
	rsBundle.CursorLocation = 2;
	rsBundle.LockType = 3;
	rsBundle.Open();
}

var rsSysOptr = Server.CreateObject("ADODB.Recordset");
rsSysOptr.ActiveConnection = MM_cnnASP02_STRING;
rsSysOptr.Source = "{call dbo.cp_SysOptr(0,0,20)}";
rsSysOptr.CursorType = 0;
rsSysOptr.CursorLocation = 2;
rsSysOptr.LockType = 3;
rsSysOptr.Open();
%>
<html>
<head>
	<title>Bundle Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m005Srh01.js"></script>
	<script language="javascript">
	if (window.focus) self.focus();
	function CnstrFltr() {
		var chvOptr = document.frm0102.BundleNameOperator[document.frm0102.BundleNameOperator.selectedIndex].value ;
		var chvStg1 = document.frm0102.BundleName.value;
		var stgFilter = ACfltr_05("122","",chvOptr,chvStg1,"");		
		stgPgQuery = "m010p0102.asp?";
		stgPgQuery   += "Search=true&inspSrtBy=1&inspSrtOrd=0&chvFilter=" + stgFilter ;
		document.frm0102.action = stgPgQuery ;
		document.frm0102.submit(); 
	}
	
	function SelectBundle(){
		if (document.frm0102.SearchResult.selectedIndex==-1){
			alert("Select a bundle.")
			return ;
		}	
		if (!top.opener.closed) {
			top.opener.document.frm0201.ClassName.value=document.frm0102.SearchResult.options[document.frm0102.SearchResult.selectedIndex].text;
			top.opener.document.frm0201.ClassID.value=document.frm0102.SearchResult[document.frm0102.SearchResult.selectedIndex].value;
			top.opener.document.frm0201.ClassBundle.value=0;
			if (document.frm0102.ListUnitCost.length > 1) {
				top.opener.document.frm0201.ListUnitCost.value=document.frm0102.ListUnitCost[document.frm0102.SearchResult.selectedIndex].value;
			} else {
				top.opener.document.frm0201.ListUnitCost.value=document.frm0102.ListUnitCost.value;			
			}
			top.opener.CalculateTotal();				
		}
		top.window.close();
	}
	
	function init(){
	<% 
	if (String(Request.QueryString("Search")) == "true") { 
	%>
		document.frm0102.SearchResult.focus();
	<% 
	} else { 
	%>
		document.frm0102.BundleNameOperator.focus();
	<% 
	} 
	%>
	}	
	</script>
</head>
<body onLoad="init();">
<form name="frm0102" method="POST" action="">
<h5>Search Criteria</h5>
<table cellpadding="1" cellspacing="1">	
	<tr>
		<td nowrap>
			Bundle Name:		
			<select name="BundleNameOperator" tabindex="1" accesskey="F">
<% 
		while (!rsSysOptr.EOF) { 
			if (rsSysOptr.Fields.Item("intRecID").Value == 122) {
%>
				<option value="<%=(rsSysOptr.Fields.Item("intOptrId").Value)%>"><%=(rsSysOptr.Fields.Item("chvOptrDesc").Value)%>
<% 	
			}
			rsSysOptr.MoveNext(); 
		}
%>
			</select>
			<input type="text" name="BundleName" size="40" tabindex="2" maxlength="50" value="<%=Request.QueryString("BundleName")%>">
			<input type="button" value="Search" onClick="CnstrFltr();" tabindex="3" class="btnstyle">
		</td>		
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><select name="SearchResult" size="20" style="width:420px" tabindex="4">
<% 
		if (String(Request.QueryString("Search")) == "true") { 
			var count = 0;
			while (!rsBundle.EOF) {
				count++;
%>
				<option value="<%=(rsBundle.Fields.Item("insBundle_id").Value)%>"><%=(rsBundle.Fields.Item("chvName").Value)%>
<%
				rsBundle.MoveNext();
			}
		}
%>
		</select></td>		
	</tr>
	<tr>
		<td><input type="button" value="Select Bundle" tabindex="5" onClick="SelectBundle();" class="btnstyle"></td>
	</tr>
</table>
<% 
if (String(Request.QueryString("Search")) == "true") { 
	if (count > 0) rsBundle.MoveFirst();
		while (!rsBundle.EOF) {
%>
<input type="hidden" name="ListUnitCost" value="<%=(rsBundle.Fields.Item("FltList_Unit_Cost").Value)%>">
<%
			rsBundle.MoveNext();
		}
	rsBundle.Close();
}
%>
</form>
</body>
</html>
<%
rsSysOptr.Close();
%>
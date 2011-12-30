<!--------------------------------------------------------------------------
* File Name: m014p0203.asp
* Title: Institution Search
* Main SP: cp_school3
* Description: This page is used to find an institution.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
// retrieve search operands
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup2(743,0,'',0,'',0)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

if (String(Request.QueryString("Search"))=="true") {
	var rsInstitution__inspSrtBy = "1";
	if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
	  rsInstitution__inspSrtBy = String(Request.QueryString("inspSrtBy"));
	}
	var rsInstitution__inspSrtOrd = "0";
	if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
	  rsInstitution__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
	}
	var rsInstitution__chvFilter = "";
	if(String(Request.QueryString("chvFilter")) != "undefined") { 
	  rsInstitution__chvFilter = String(Request.QueryString("chvFilter"));
	}
	
	var rsInstitution = Server.CreateObject("ADODB.Recordset");
	rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitution.Source = "{call dbo.cp_school3(0,'',0,0,0,0,0,"+rsInstitution__inspSrtBy+","+rsInstitution__inspSrtOrd+",'"+rsInstitution__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
	rsInstitution.CursorType = 0;
	rsInstitution.CursorLocation = 2;
	rsInstitution.LockType = 3;
	rsInstitution.Open();
}
%>
<html>
<head>
	<title>Institution Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m012Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>	
	<script language="JavaScript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,27)}";
	rsOptr.CursorType = 0;
	rsOptr.CursorLocation = 2;
	rsOptr.LockType = 3;
	rsOptr.Open();
	
	if (!rsOptr.EOF){ 	
		while (!rsOptr.EOF) { 
			objOptrDesc = rsOptr("chvOptrDesc")
			objOptrId   = rsOptr("intOptrId")
			objRecID    = rsOptr("intRecID")	
			if (intOldOptr != objRecID.value) {
				Response.Write("detailData["+objRecID+"] = new Array();")
				intOldOptr = objRecID.value
			}
			Response.Write("detailData["+objRecID+"]["+objOptrId+"] = '"+ objOptrDesc+"';");
			rsOptr.MoveNext 
		}
	}
	else {
	   Response.Write("SysOptr lookup does not exist.")
	}	
	rsOptr.Close();
	%>

	function Init() {
		selectChange(frm14p03.Operand, frm14p03.Operator,detailData);
	<%
	if (String(Request.QueryString("Search")) == "true") { 		
	%>
		document.frm14p03.SearchResult.focus();
	<%
	} else {
	%>
		document.frm14p03.Operand.focus();
	<%
	}
	%>		
	}

	function selectChange(control, controlToPopulate,ItemArray) {
		document.frm14p03.StringSearchTextBoxOne.value="";
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = control.value;
		if (y != 0) {
			for ( x = 0 ; x < ItemArray[y].length  ; x++ ) {
				if (ItemArray[y][x]) { 
					myEle = document.createElement("option") ;
					myEle.value = x ;
					myEle.text = ItemArray[y][x] ;
					controlToPopulate.add(myEle) ;
			   }
			}
		}
		Togo();	  
	}
		
	function Togo() {
		var j = document.frm14p03.Operand[document.frm14p03.Operand.selectedIndex].value ;
		var l = document.frm14p03.Operator[document.frm14p03.Operator.selectedIndex].value	
		document.frm14p03.MM_curOprd.value = j ;
		document.frm14p03.MM_curOptr.value = l ;
		document.frm14p03.MM_flag.value = true ;
	}

	function CnstrFltr() {
		var stgPgQuery = "";
		
		var chvOprd = document.frm14p03.Operand[document.frm14p03.Operand.selectedIndex].value ; 
		var chrNot  = "";
		var chvOptr = document.frm14p03.Operator[document.frm14p03.Operator.selectedIndex].value ;
		var chvStg1 = document.frm14p03.StringSearchTextBoxOne.value;
		var chvStg2 = "";
		
		if (chvOptr == "0") {
			document.write("<B><B><B>");
			document.write("Please select operator before Proceed");
			document.write("<B><B><B>");
		} else {	
			var stgFilter = ACfltr_12(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
			stgPgQuery += "m014p0203.asp?Search=true&inspSrtBy=1&inspSrtOrd=0&chvFilter=" + stgFilter ;
		}
		document.frm14p03.action = stgPgQuery;
		document.frm14p03.submit();
	}
	
	function SelectInstitution(){
		if (document.frm14p03.SearchResult.selectedIndex==-1){
			alert("Select an institution.")
			return ;
		}
		opener.document.frm0201.UserID.value=document.frm14p03.SearchResult[document.frm14p03.SearchResult.selectedIndex].value;
		opener.document.frm0201.UserName.value=document.frm14p03.SearchResult.options[document.frm14p03.SearchResult.selectedIndex].text;		
		opener.document.frm0201.UserType.value="1";
		self.close();
	}
	</script>
</head>
<body onload="Init();">
<form name="frm14p03" method="post" action="">
<h5>Search Criteria</h5>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>
			<select name="Operand" onchange="selectChange(this, frm14p03.Operator,detailData);" tabindex="1" accesskey="F">
		<% 
		while (!rsOprd.EOF) {
			if (rsOprd.Fields.Item("intRecID").Value==135) {
		%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
		<%
			}
			rsOprd.MoveNext();
		}
		%>
			</select>
			<select name="Operator" onchange="Togo();" tabindex="3" style="width:130px"></select>
			<input type="text" name="StringSearchTextBoxOne" tabindex="4">
		</td>
	</tr>
	<tr>
		<td nowrap colspan="5"><input type="button" value="Proceed" onClick="CnstrFltr();" tabindex="6" class="btnstyle"></td>		
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><select name="SearchResult" size="20" style="width:420px" tabindex="7">
		<% 
		var count = 0;
		if (String(Request.QueryString("Search")) == "true") { 
			while (!rsInstitution.EOF) {
		%>
				<option value="<%=(rsInstitution.Fields.Item("insSchool_id").Value)%>"><%=(rsInstitution.Fields.Item("chvSchool_Name").Value)%>
		<%
				rsInstitution.MoveNext();
				count++;			
			}
		}
		%>
		</select></td>		
	</tr>
	<tr>
		<td>
			<input type="button" value="Select Institution" tabindex="8" onClick="SelectInstitution();" <%=((count==0)?"DISABLED":"")%> class="btnstyle">
			<input type="button" value="Close" tabindex="9" onClick="top.window.close();" class="btnstyle">
		</td>		
	</tr>
</table>
<input type="hidden" name="MM_flag" value="false">
<input type="hidden" name="MM_curOprd">
<input type="hidden" name="MM_curOptr">
</form>
</body>
</html>
<%
rsOprd.Close();
%>
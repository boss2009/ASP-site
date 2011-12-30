<!--------------------------------------------------------------------------
* File Name: m014p0401.asp
* Title: PILAT Student Search
* Main SP: cp_pilat_student
* Description: This page is used to find a client.
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
rsOprd.Source = "{call dbo.cp_asP_lkup2(747,0,'',0,'',0)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

if (String(Request.QueryString("Search"))=="true") {
	var rsPILATStudent__inspSrtBy = "1";
	if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
		rsPILATStudent__inspSrtBy = String(Request.QueryString("inspSrtBy"));
	}
	var rsPILATStudent__inspSrtOrd = "0";
	if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
		rsPILATStudent__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
	}
	var rsPILATStudent__chvFilter = "";
	if(String(Request.QueryString("chvFilter")) != "undefined") { 
		rsPILATStudent__chvFilter = String(Request.QueryString("chvFilter"));
	}
	
	var rsPILATStudent = Server.CreateObject("ADODB.Recordset");
	rsPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
	rsPILATStudent.Source = "{call dbo.cp_pilat_student(0,'','','','','',0,0,0,0,'',0,0,0,"+rsPILATStudent__inspSrtBy+","+rsPILATStudent__inspSrtOrd+",'"+rsPILATStudent__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
	rsPILATStudent.CursorType = 0;
	rsPILATStudent.CursorLocation = 2;
	rsPILATStudent.LockType = 3;
	rsPILATStudent.Open();
}

if (String(Request.QueryString("Save"))=="true") {
	var rsAddPILATStudent = Server.CreateObject("ADODB.Recordset");
	rsAddPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
	rsAddPILATStudent.Source = "{call dbo.cp_pilat_stdsch_ref(0,"+Request.QueryString("intReferral_id")+","+Request.QueryString("intPStdnt_id")+",0,'A',0)}";
	rsAddPILATStudent.CursorType = 0;
	rsAddPILATStudent.CursorLocation = 2;
	rsAddPILATStudent.LockType = 3;
	rsAddPILATStudent.Open();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>PILAT Student Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m022Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>	
	<script language="JavaScript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var intIdxOprd = 0;
	var objOptrDesc,objOptrId,objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,34)}";
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
				intIdxOprd += 1
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
		selectChange(frm12p04.Operand, frm12p04.Operator,detailData);
	<%
	if (String(Request.QueryString("Search")) == "true") { 		
	%>
		document.frm12p04.SearchResult.focus();
	<%
	} else {
	%>
		document.frm12p04.Operand.focus();
	<%
	}
	%>		
	}

	function selectChange(control, controlToPopulate,ItemArray) {
		document.frm12p04.StringSearchTextBoxOne.value="";
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
		var j = document.frm12p04.Operand[document.frm12p04.Operand.selectedIndex].value ;
		var l = document.frm12p04.Operator[document.frm12p04.Operator.selectedIndex].value	
		document.frm12p04.MM_curOprd.value = j ;
		document.frm12p04.MM_curOptr.value = l ;
		document.frm12p04.MM_flag.value = true ;
	}

	function CnstrFltr() {
		var stgPgQuery = "";
		
		var chvOprd = document.frm12p04.Operand[document.frm12p04.Operand.selectedIndex].value ; 
		var chrNot  = "";
		var chvOptr = document.frm12p04.Operator[document.frm12p04.Operator.selectedIndex].value ;
		var chvStg1 = document.frm12p04.StringSearchTextBoxOne.value;
	
		if (chvOprd == "158") {
			if (!IsID(chvStg1)) {
				alert("Invalid number.");
				document.frm08p02.StringSearchTextBoxOne.focus();
				return;			
			}
		}		
		
		if (chvOptr == "0") {
			document.write("<B><B><B>");
			document.write("Please select operator before Proceed");
			document.write("<B><B><B>");
		} else {	
			var stgFilter = ACfltr_22(chvOprd,chrNot,chvOptr,chvStg1,"");
			stgPgQuery += "m012p0401.asp?Search=true&intReferral_id=<%=Request.QueryString("intReferral_id")%>&inspSrtBy=1&inspSrtOrd=0&chvFilter=" + stgFilter ;
		}
		document.frm12p04.action = stgPgQuery;
		document.frm12p04.submit();
	}
	
	function SelectStudent(){
		if (document.frm12p04.SearchResult.selectedIndex==-1){
			alert("Select a PILAT student.")
			return ;
		}
		document.frm12p04.action = "m012p0401.asp?Save=true&intPStdnt_id="+document.frm12p04.SearchResult[document.frm12p04.SearchResult.selectedIndex].value+"&intReferral_id=<%=Request.QueryString("intReferral_id")%>";
		document.frm12p04.submit();	
	}
	</script>
</head>
<body onload="Init();">
<form name="frm12p04" method="post" action="">
<h5>Search Criteria</h5>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>
			<select name="Operand" onchange="selectChange(this, frm12p04.Operator,detailData);" tabindex="1" accesskey="F">
			<% 
			while (!rsOprd.EOF) {
			%>
				<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value == "163")?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
			<%
				rsOprd.MoveNext();
			}
			%>
			</select>
			<select name="Operator" onchange="Togo();" tabindex="3"></select>
			<input type="text" name="StringSearchTextBoxOne" tabindex="4" size="15">
			<input type="button" value="Proceed" onClick="CnstrFltr();" tabindex="6" class="btnstyle">
		</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><select name="SearchResult" size="20" style="width:420px" tabindex="7">
		<% 
		var count = 0;
		if (String(Request.QueryString("Search")) == "true") { 
			while (!rsPILATStudent.EOF) {
		%>
				<option value="<%=(rsPILATStudent.Fields.Item("intPStdnt_id").Value)%>"><%=(rsPILATStudent.Fields.Item("chvLst_Name").Value)%>, <%=(rsPILATStudent.Fields.Item("chvFst_Name").Value)%>
		<%
			rsPILATStudent.MoveNext();
			count++;			
			}
		}
		%>
		</select></td>		
	</tr>
	<tr>
		<td>
			<input type="button" value="Select Student" tabindex="8" onClick="SelectStudent();" <%=((count==0)?"DISABLED":"")%> class="btnstyle">
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
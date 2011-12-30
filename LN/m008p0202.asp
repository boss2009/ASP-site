<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
// retrieve search operands
var rsOprd = Server.CreateObject("ADODB.Recordset");
rsOprd.ActiveConnection = MM_cnnASP02_STRING;
rsOprd.Source = "{call dbo.cp_ASP_Lkup(715)}";
rsOprd.CursorType = 0;
rsOprd.CursorLocation = 2;
rsOprd.LockType = 3;
rsOprd.Open();

if (String(Request.QueryString("Search"))=="true") {
	var rsClient__inspSrtBy = "1";
	var rsClient__inspSrtOrd = "0";
	var rsClient__chvFilter = "";		
	if (String(Request.QueryString("inspSrtBy")) != "undefined") rsClient__inspSrtBy = String(Request.QueryString("inspSrtBy"));
	if (String(Request.QueryString("inspSrtOrd")) != "undefined") rsClient__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
	if(String(Request.QueryString("chvFilter")) != "undefined") rsClient__chvFilter = String(Request.QueryString("chvFilter"));
	
	var rsClient = Server.CreateObject("ADODB.Recordset");
	rsClient.ActiveConnection = MM_cnnASP02_STRING;
	rsClient.Source = "{call dbo.cp_Adult_Client("+ rsClient__inspSrtBy.replace(/'/g, "''") + ","+ rsClient__inspSrtOrd.replace(/'/g, "''") + ",'"+ rsClient__chvFilter.replace(/'/g, "''") + "')}";
	rsClient.CursorType = 0;
	rsClient.CursorLocation = 2;
	rsClient.LockType = 3;
	rsClient.Open();
}
%>
<html>
<head>
	<title>Client Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m001Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>	
	<script language="JavaScript">
	var detailData = new Array();
	<%
	var intOldOptr = 0;
	var objRecID;
	
	var rsOptr = Server.CreateObject("ADODB.Recordset");
	rsOptr.ActiveConnection = MM_cnnASP02_STRING;
	rsOptr.Source = "{call dbo.cp_SysOptr(0,0,16)}";
	rsOptr.CursorType = 0;
	rsOptr.CursorLocation = 2;
	rsOptr.LockType = 3;
	rsOptr.Open();
	if (!rsOptr.EOF){ 	
		while (!rsOptr.EOF) { 
			objRecID    = rsOptr("intRecID")	
			if (intOldOptr != objRecID.value) {
				Response.Write("detailData["+objRecID+"] = new Array();")
				intOldOptr = objRecID.value
			}
			Response.Write("detailData["+objRecID+"]["+rsOptr("intOptrId")+"] = '"+ rsOptr("chvOptrDesc")+"';");
			rsOptr.MoveNext 
		}
	} else {
	   Response.Write("SysOptr lookup does not exist.")
	}	
	rsOptr.Close();
	%>

	function Init() {
		oStg2.style.visibility = "hidden";
		selectChange(frm08p02.Operand, frm08p02.Operator,detailData)
	}

	function selectChange(control, controlToPopulate,ItemArray) {
		document.frm08p02.StringSearchTextBoxOne.value="";
		document.frm08p02.StringSearchTextBoxTwo.value="";
		var myEle ;
		for (var q=controlToPopulate.options.length;q>=0;q--) controlToPopulate.options[q]=null;
		myEle = document.createElement("option") ;
		var y = control.value;
		if ((y == "18") || (y == "22")){
			document.frm08p02.StringSearchTextBoxTwo.disabled = false;
			oStg2.style.visibility = "visible";
		} else {
			document.frm08p02.StringSearchTextBoxTwo.disabled = true;
			oStg2.style.visibility = "hidden";
		}
		if (y == "17") {
			oStg1.style.visibility="hidden";
		} else {
			oStg1.style.visibility="visible";
		}
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
		var j = document.frm08p02.Operand[document.frm08p02.Operand.selectedIndex].value ;
		var l = document.frm08p02.Operator[document.frm08p02.Operator.selectedIndex].value	
		document.frm08p02.MM_curOprd.value = j ;
		document.frm08p02.MM_curOptr.value = l ;
		document.frm08p02.MM_flag.value = true ;
	}

	function CnstrFltr() {
		var stgPgQuery = "";
		
		var chvOprd = document.frm08p02.Operand[document.frm08p02.Operand.selectedIndex].value ; 
		var chrNot  = "";
		var chvOptr = document.frm08p02.Operator[document.frm08p02.Operator.selectedIndex].value ;
		var chvStg1 = document.frm08p02.StringSearchTextBoxOne.value;
		var chvStg2 = document.frm08p02.StringSearchTextBoxTwo.value;

		if (((chvOprd=="18") || (chvOprd=="22")) && (chvOptr!="0")) {
			if (chvStg1 == "") {
				alert("Enter Start Date");
				document.frm08p02.StringSearchTextBoxOne.focus();
				return ;
			}
			if (chvStg2 == "") {
				alert("Enter End Date");
				document.frm08p02.StringSearchTextBoxTwo.focus();
				return ;
			}
			if (!CheckDateBetween(Trim(chvStg1)+" "+Trim(chvStg2))) {
				return ;
			}
		}

		if (chvOprd == "33") {
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
			var stgFilter = ACfltr_01(chvOprd,chrNot,chvOptr,chvStg1,chvStg2);
			stgPgQuery += "m008p0202.asp?Search=true&inspSrtBy=1&inspSrtOrd=0&chvFilter=" + stgFilter ;
		}
		document.frm08p02.action = stgPgQuery;
		document.frm08p02.submit();
	}
	
	function SelectClient(){
		if (document.frm08p02.SearchResult.selectedIndex==-1){
			alert("Select a client.")
			return ;
		}
	
		opener.document.frm0101.UserType.value="3";
		opener.document.frm0101.IndividualUserID.value=document.frm08p02.SearchResult[document.frm08p02.SearchResult.selectedIndex].value;
		opener.document.frm0101.IndividualUserName.value=document.frm08p02.SearchResult.options[document.frm08p02.SearchResult.selectedIndex].text;
		self.close();
	}
	</script>
</head>
<body onload="Init();">
<form name="frm08p02" method="post" action="">
<h5>Search Criteria</h5>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap><select name="Operand" onchange="selectChange(this, frm08p02.Operator,detailData);" tabindex="1" accesskey="F">
		<% 
		while (!rsOprd.EOF) {
		%>
			<option value="<%=(rsOprd.Fields.Item("intRecID").Value)%>" <%=((rsOprd.Fields.Item("intRecID").Value == Request.Form("MM_curOprd"))?"SELECTED":"")%> ><%=(rsOprd.Fields.Item("chvObjName").Value)%></option>
		<%
			rsOprd.MoveNext();
		}
		%>
		</select></td>
		<td nowrap><select name="Operator" onchange="Togo();" tabindex="3" style="width:130px"></select></td>
		<td nowrap><DIV ID="oStg1" STYLE="visibility:visible">
			<input type="text" name="StringSearchTextBoxOne" tabindex="4" size="10">
		</DIV></td>
		<td nowrap><DIV ID="oStg2" STYLE="visibility:hidden">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
			<input type="text" name="StringSearchTextBoxTwo" tabindex="5" accesskey="L" size="10" DISABLED>
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</DIV></td>
	</tr>
	<tr>
		<td nowrap colspan="5"><input type="button" value="Search" onClick="CnstrFltr();" tabindex="6" class="btnstyle"></td>		
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><select name="SearchResult" size="20" style="width:420px" tabindex="7">
		<% 
		var count = 0;
		if (String(Request.QueryString("Search")) == "true") { 
			while (!rsClient.EOF) {
		%>
				<option value="<%=(rsClient.Fields.Item("intAdult_Id").Value)%>"><%=(rsClient.Fields.Item("chvLst_Name").Value)%>, <%=(rsClient.Fields.Item("chvFst_Name").Value)%>
		<%
				rsClient.MoveNext();
				count++;			
			}
		}
		%>
		</select></td>		
	</tr>
	<tr>
		<td>
			<input type="button" value="Select Client" tabindex="8" onClick="SelectClient();" <%=((count==0)?"DISABLED":"")%> class="btnstyle">
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
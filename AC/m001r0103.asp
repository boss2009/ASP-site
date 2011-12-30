<%@language="JAVASCRIPT"%>
<!--#include virtual="ASP/inc/ASPUtility.inc" -->
<!--#include virtual="ASP/inc/ASPCheckLogin.inc" -->
<!--#include virtual="ASP/Connections/cnnASP02.asp" -->
<% 
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup(714)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

// retrieve only the B-code service type    + Jun.16.2004 
var rsServiceType = Server.CreateObject("ADODB.Recordset");
rsServiceType.ActiveConnection = MM_cnnASP02_STRING;
rsServiceType.Source = "{call dbo.cp_service_type(0,0,0,3)}";
rsServiceType.CursorType = 0;
rsServiceType.CursorLocation = 2;
rsServiceType.LockType = 3;
rsServiceType.Open();	
%>
<html>
<head>
	<title>Service, Date Range Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m001Srh04.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="javascript">
// -----------------------------------------
// function CnstrFltr()
// -----------------------------------------
	function CnstrFltr(output) {		
	
//		var inspSrtBy = document.frm01r0103.SortByColumn.value;
//		var inspSrtOrd = document.frm01r0103.OrderBy.value;

// Sort by Service type followed by Adult Client is a must     + Jun.23.2004  
		var inspSrtBy = 11;
		var inspSrtOrd = 0;

		var stgFilter = document.frm01r0103.QueryString.value;
//		var Show = document.frm01r0103.Show.value;
		if (output==1) {

//			document.frm01r0103.action = "m001r0103q.asp?Show="+Show+"&inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;
			document.frm01r0103.action = "m001r0103q.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;
			document.frm01r0103.submit();
		} else {
			var SearchExcel = window.open("m001r0103excel.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter);
		}
	}
// -----------------------------------------
// function initscr()
// -----------------------------------------
	function initscr() {
		oStg1.style.visibility="visible";
		oStg2.style.visibility="hidden";
		oStg3.style.visibility="hidden";
		initializeList('frm01r0103.MultiValueServiceType','frm01r0103.MultiValueSearchOptions');	
		document.frm01r0103.SearchType[0].focus();		
	}

	function SelOpt() {
		var len = document.frm01r0103.SearchType.length;
		var Idparam = 1;
		for (var i=0;i <len; i++){
			if (document.frm01r0103.SearchType[i].checked)  Idparam = i;
		}
	
		switch ( Idparam ) {
			case 0: 
				oStg1.style.visibility="visible";
				oStg2.style.visibility="hidden";
				oStg3.style.visibility="hidden";
			break;
			case 1: 
				oStg1.style.visibility="hidden";
				oStg2.style.visibility="visible";
				oStg3.style.visibility="visible";
			break;
		}       	
	}

	function Savtxt() {	
		var stgPgQuery = "";
		var stgFilter = "" ;
		var blnFlg = false ;
		var len = document.frm01r0103.SearchType.length;
		var Idparam = 1;	
		var stgTemp,j,k; 
	  	
		for (var i=0;i <len; i++){
			if (document.frm01r0103.SearchType[i].checked) Idparam = i;
		}
	
		stgTemp = document.frm01r0103.QueryString.value; 
		switch (Idparam) {
			case 0: 
				var chvStg1 = document.frm01r0103.StringSearchTextOne.value;
				var chvStg2 = document.frm01r0103.StringSearchTextTwo.value;
				var chvAO1  = document.frm01r0103.AndOr.value ;
				if (chvStg1 == "") {
					alert("Enter Start Date");
					document.frm01r0103.StringSearchTextOne.focus();
					return ;
				}
				if (chvStg2 == "") {
					alert("Enter End Date");
					document.frm01r0103.StringSearchTextTwo.focus();
					return ;
				}
				if (!CheckDateBetween(Trim(chvStg1)+" and "+Trim(chvStg2))) {
					return ;
				}
				stgFilter = ACfltr_03("14","","",chvStg1,chvStg2);
				if (stgTemp.length > 0 ) stgTemp += " (" ;
				stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;
			break;
			case 1: 
				l = "";
				var count = 0;
				var optList = document.frm01r0103.MultiValueSearchOptions;
				for (var x= 0; x < optList.length; x++){
					if (optList[x].selected) count++;				
				}	
				if (count == 0) {
					alert("Select at least one value.");					
					optList.focus();
					return ;	
				}								
				var m = optList.length;
				if (optList.multiple) {
					for(var ii = 0; ii < m; ii++) {
						if (document.frm01r0103.MultiValueSearchOptions[ii].selected) {		
							if (l.length > 0 ) l += "," ;
							l += document.frm01r0103.MultiValueSearchOptions[ii].value ;		
						} 				  
					} 						  
				} else {
					l = document.frm01r0103.MultiValueSearchOptions[document.frm01r0103.MultiValueSearchOptions.selectedIndex].value
				} 
				if (stgTemp.length > 0 ) stgTemp += " (" ;
				blnFlg = true;				
				stgFilter += " insSrv_Code_id IN (" + l + ") " ; 
			break;
			default: 
			break;
		}						
		document.frm01r0103.StringSearchTextOne.value="";
		if (blnFlg) {
			var chvAO1  = document.frm01r0103.AndOr.value ;
			stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 + " ";
		}
		document.frm01r0103.QueryString.value = stgTemp; 
	}
/////////////////////////////////////////////
// -----------------------------------------
// function Tip_on()
// -----------------------------------------
function Tip_on(stgchar){
//alert("Tip_on - "+stgchar);
  switch (stgchar) {
     case "1" :  
	             TheTip1.style.visibility="visible";
                 break;  
     case "2" :  
	             TheTip2.style.visibility="visible";
                 break;  
     case "3" :  
	             TheTip3.style.visibility="visible";
                 break;  
  }
}
// -----------------------------------------
// function Tip_off()
// -----------------------------------------
function Tip_off(stgchar) {
//alert("Tip_off - "+stgchar);
  switch (stgchar) {
     case "1" :  
//	             TheTip1.innerHTML = "";
//               TheTip1.outerHTML = "";
	             TheTip1.style.visibility="hidden";
                 break;  
     case "2" :  
//	             TheTip2.innerHTML = "";
//                TheTip2.outerHTML = "";
	             TheTip2.style.visibility="hidden";
                 break;  
     case "3" :  
	             TheTip3.style.visibility="hidden";
                 break;  
  }
}
////////////////////////////////////////////
var objParent ;
var objChild  ;
// -----------------------------------------
// function initializeList to populate list.
// -----------------------------------------
function initializeList(parentList, childList){
   objParent = eval("document." + parentList);
   objChild = eval("document." + childList);
   if (objParent.selectedIndex < 0) {
	   objParent.options[0].selected = true
   }
   var OptnID = objParent.options[objParent.selectedIndex].value;	
// map the Operand ID to the detailArray
// OptnID -= 26 ;

//   var hfrm = document.frm01r0103 ;
//       oOld2.style.visibility="hidden";
//       hfrm.ckbOld.disabled = true;
//
//alert("OptnID - "+OptnID);
clearList(objChild)
fillDetail(OptnID)
}
// -----------------------------------------
// function fillDetail()
// -----------------------------------------
function fillDetail(deptID){
if(detailData2[deptID]){
	for(x=0; x < detailData2[deptID].length; x++){
		if (detailData2[deptID][x]) {
			var sDesc = detailData2[deptID][x];
			var valOption = new Option(sDesc);
	 		valOption.value = x;
	 		objChild.options[objChild.length] = valOption;
		}
	}
	if (objChild.length) {
		objChild.options[0].selected = true;
	}
}
objChild = null;
objParent = null;

}
// -----------------------------------------
// function clearList()
// -----------------------------------------
function clearList(obj){
	if (obj.length){obj.options.length = 0;}
}
	</script>

	<script language="javascript">
// -----------------------------------------
// Build the List Box Array which house multiple lookups content
// -----------------------------------------
   var detailData2 = new Array();
// EPPD is Group 1
// ------------------------ set up parameters for EPPD Service Type lookup
<%
// server side variables
   var intEPPDCnt = 0;
   var intSSBCnt  = 0;
   var intID,chrData;	
// -----------------------------------------
// EPPD Service Type lookup 
// -----------------------------------------
	var rsServiceTypeEPPD = Server.CreateObject("ADODB.Recordset");
	rsServiceTypeEPPD.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceTypeEPPD.Source = "{call dbo.cp_service_type(0,0,0,3)}";
	rsServiceTypeEPPD.CursorType = 0;
	rsServiceTypeEPPD.CursorLocation = 2;
	rsServiceTypeEPPD.LockType = 3;
	rsServiceTypeEPPD.Open();	
	if (!rsServiceTypeEPPD.EOF){ 
		Response.Write("detailData2[1] = new Array();")	
		while (!rsServiceTypeEPPD.EOF) { 
			Response.Write("detailData2[1]["+rsServiceTypeEPPD("insService_type_id")+"] = '"+rsServiceTypeEPPD("chvname")+"';");
			rsServiceTypeEPPD.MoveNext 
		}
	} else {
		Response.Write("EPPD Service Code lookup does not exist.")
	}
	rsServiceTypeEPPD.Close();
	
	var rsServiceTypeSSB = Server.CreateObject("ADODB.Recordset");
	rsServiceTypeSSB.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceTypeSSB.Source = "{call dbo.cp_service_type(0,0,0,4)}";
	rsServiceTypeSSB.CursorType = 0;
	rsServiceTypeSSB.CursorLocation = 2;
	rsServiceTypeSSB.LockType = 3;
	rsServiceTypeSSB.Open();	
	if (!rsServiceTypeSSB.EOF){ 
		Response.Write("detailData2[2] = new Array();")	
		while (!rsServiceTypeSSB.EOF) { 
			Response.Write("detailData2[2]["+rsServiceTypeSSB("insService_type_id")+"] = '"+rsServiceTypeSSB("chvname")+"';");
			rsServiceTypeSSB.MoveNext 
		}
	} else {
		Response.Write("SSB Service Code lookup does not exist.")
	}
	rsServiceTypeSSB.Close();

%>
	</script>

</head>
<body onload="initscr();">
<form name="frm01r0103" method="post" action="">
  <h5>B-Code Service, Date Range Report</h5>
<hr>
<table cellpadding="1" cellspacing="1" border="0">
    <tr>
		<td nowrap valign="top" width="120"><input type="radio" name="SearchType" value="1" checked onClick="SelOpt()" accesskey="F" tabindex="1" class="chkstyle">Service Date</td>
		<td nowrap valign="center"><DIV ID="oStg1" STYLE="visibility:visible">
			Between
			<input type="text" name="StringSearchTextOne" tabindex="2" size="11" maxlength="10">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
			<input type="text" name="StringSearchTextTwo" value="<%=CurrentDate()%>" tabindex="3" size="11" maxlength="10">
           	<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</DIV></td>
    </tr>
</table>
  <table cellpadding="1" cellspacing="1" border="0" width="448">
    <tr> 
		
      <td nowrap valign="top" width="117"> 
        <input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="4" class="chkstyle">Service Type</td>
      <td nowrap valign="top" width="65"> 
        <DIV ID="oStg3" STYLE="visibility:hidden">
           <select name="MultiValueServiceType" OnChange="initializeList('frm01r0103.MultiValueServiceType','frm01r0103.MultiValueSearchOptions')" tabindex="11">
             <option value="1">EPPD</option>
             <option value="2">SSB</option>
          </select>
		</DIV></td>
      <td nowrap valign="center" width="256"> 
        <DIV ID="oStg2" STYLE="visibility:hidden">
			Contains: 
			<select name="MultiValueSearchOptions" size="8" width="150" align="top" multiple tabindex="5">
			<%			
			while (!rsServiceType.EOF) { 
			%>
				<option value="<%=rsServiceType.Fields.Item("insService_type_id").Value%>"><%=rsServiceType.Fields.Item("chvname").Value%>
			<%
				rsServiceType.MoveNext 
			}
			%>			
			</select>
		</DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td valign="top"><select name="AndOr" tabindex="6">
			<option value=" ">None</option>
			<option value="And">And</option>
			<option value="Or">Or</option>
		</select></td>	
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top" nowrap><input type="button" name="AddToQueryString" value="Add" tabindex="7" onClick="Savtxt()" class="btnstyle"></td>
		<td valign="top" nowrap><textarea name="QueryString" cols="80" rows="6" tabindex="8" accesskey="L"></textarea></td>
    </tr>
</table>
<br>
  <table cellpadding="1" cellspacing="1" width="205">
<!--
// + Jun.17.2004
    <tr>
		<td nowrap>
			Sort by:
			<select name="SortByColumn" tabindex="9">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
        	<select name="OrderBy" tabindex="10">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
      </td>
	</tr>
-->
	<tr>
		<td nowrap>
	        <input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="12" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="13" class="btnstyle">
			<input type="button" value="Clear All" onClick="window.location.reload();" tabindex="14" class="btnstyle">
		</td>		
    </tr>
</table>
</form>
</body>
</html>
<%
rsServiceType.Close();
rsCol.Close();
%>
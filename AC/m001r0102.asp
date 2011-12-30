<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% 
// retrieve the sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup(714)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_funding_source_attributes(0,0,0,0,1,0,0,0,2,'Q',0)}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();	

var rsServiceType = Server.CreateObject("ADODB.Recordset");
rsServiceType.ActiveConnection = MM_cnnASP02_STRING;
rsServiceType.Source = "{call dbo.cp_service_type(0,0,0,0)}";
rsServiceType.CursorType = 0;
rsServiceType.CursorLocation = 2;
rsServiceType.LockType = 3;
rsServiceType.Open();	
%>
<html>
<head>
	<title>Funding Source, Service, Date Range Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m001Srh04.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="javascript">
	function CnstrFltr(output) {		
		var stgFilter = document.frm01s01.QueryString.value;
		if (output==1) {
			document.frm01s01.action = "m001r0102q.asp?chvFilter=" + stgFilter ;
			document.frm01s01.submit() ; 
		} else {
			var SearchExcel = window.open("m001r0102excel.asp?chvFilter=" + stgFilter);			
		}
	}

	function initscr() {
		oStg1.style.visibility="visible";
		oStg2.style.visibility="hidden";
		oStg3.style.visibility="hidden";
		document.frm01s01.SearchType[0].focus();		
	}

	function SelOpt() {
		var len = document.frm01s01.SearchType.length;
		var Idparam = 1;
		for (var i=0;i <len; i++){
			if (document.frm01s01.SearchType[i].checked)  Idparam = i;
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
				oStg3.style.visibility="hidden";
			break;
			case 2: 
				oStg1.style.visibility="hidden";
				oStg2.style.visibility="hidden";
				oStg3.style.visibility="visible";
			break;
		}       	
	}

	function Savtxt() {	
		var stgPgQuery = "";
		var stgFilter = "" ;
		var blnFlg = false ;
		var len = document.frm01s01.SearchType.length;
		var Idparam = 1;	
		var stgTemp,j,k; 
	  	
		for (var i=0;i <len; i++){
			if (document.frm01s01.SearchType[i].checked) Idparam = i;
		}
	
		stgTemp = document.frm01s01.QueryString.value; 
		switch (Idparam) {
			case 0: 
				var chvStg1 = document.frm01s01.StringSearchTextOne.value;
				var chvStg2 = document.frm01s01.StringSearchTextTwo.value;
				var chvAO1  = document.frm01s01.AndOr.value ;
				if (chvStg1 == "") {
					alert("Enter Start Date");
					document.frm01s01.StringSearchTextOne.focus();
					return ;
				}
				if (chvStg2 == "") {
					alert("Enter End Date");
					document.frm01s01.StringSearchTextTwo.focus();
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
				var chvAO1  = document.frm01s01.AndOr.value ;
				var chvStg1 = document.frm01s01.LookupValueSearchOptions.value;
				stgFilter = ACfltr_03("31","","",chvStg1,"");
				if (stgTemp.length > 0 ) stgTemp += " (" ;
				stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 ;				
			break;			
			case 2: 
				l = "";
				var count = 0;
				var optList = document.frm01s01.MultiValueSearchText;
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
						if (document.frm01s01.MultiValueSearchText[ii].selected) {		
							if (l.length > 0 ) l += "," ;
							l += document.frm01s01.MultiValueSearchText[ii].value ;		
						} 				  
					} 						  
				} else {
					l = document.frm01s01.MultiValueSearchText[document.frm01s01.MultiValueSearchText.selectedIndex].value
				} 
				if (stgTemp.length > 0 ) stgTemp += " (" ;
				blnFlg = true;				
				stgFilter += " insSrv_Code_id IN (" + l + ") " ; 
			break;
			default: 
			break;
		}						
		document.frm01s01.StringSearchTextOne.value="";
		if (blnFlg) {
			var chvAO1  = document.frm01s01.AndOr.value ;
			stgTemp +=  stgFilter  + ((stgTemp.length > 0) ? ") " : " ") + chvAO1 + " ";
		}
		document.frm01s01.QueryString.value = stgTemp; 
	}
	</script>
</head>
<body onload="initscr();">
<form name="frm01s01" method="post">
<h5>Funding Source, Service, Date Range Summary Report</h5>
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
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top" width="120"><input type="radio" name="SearchType" value="2" onClick="SelOpt()" tabindex="4" class="chkstyle">Funding Source</td>
		<td nowrap valign="center"><DIV ID="oStg2" STYLE="visibility:visible"> 
			Is 
			<select name="LookupValueSearchOptions" tabindex="5">
			<%
			while (!rsFundingSource.EOF){ 
			%>
				<option value="<%=rsFundingSource("insFunding_source_id")%>"><%=rsFundingSource.Fields.Item("chvfunding_source_name").Value%>
			<%
				rsFundingSource.MoveNext 
			}
			%>
			</select>
        </DIV></td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1" border="0">
    <tr> 
		<td nowrap valign="top" width="120"><input type="radio" name="SearchType" value="3" onClick="SelOpt()" tabindex="6" class="chkstyle">Service Type</td>
		<td nowrap valign="center"><DIV ID="oStg3" STYLE="visibility:hidden">
			Contains: 
			<select name="MultiValueSearchText" size="8" width="150" align="top" multiple tabindex="7">
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
		<td valign="top"><select name="AndOr" tabindex="8">
			<option value=" ">None</option>
			<option value="And">And</option>
			<option value="Or">Or</option>
		</select></td>	
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top" nowrap><input type="button" name="AddToQueryString" value="Add" tabindex="9" onClick="Savtxt()" class="btnstyle"></td>
		<td valign="top" nowrap><textarea name="QueryString" cols="80" rows="6" tabindex="10" accesskey="L"></textarea></td>
    </tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>
	        <input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="11" class="btnstyle">
			<input type="button" value="Excel" onClick="CnstrFltr(2);" tabindex="12" disabled class="btnstyle">
			<input type="button" value="Clear All" onClick="window.location.reload();" tabindex="13" class="btnstyle">
		</td>		
    </tr>
</table>
</form>
</body>
</html>
<%
rsFundingSource.Close();
rsServiceType.Close();
rsCol.Close();
%>
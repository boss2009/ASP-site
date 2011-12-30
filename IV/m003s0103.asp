<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup(722)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

var rsInventoryStatus = Server.CreateObject("ADODB.Recordset");
rsInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryStatus.Source = "{call dbo.cp_ASP_Lkup(36)}";
rsInventoryStatus.CursorType = 0;
rsInventoryStatus.CursorLocation = 2;
rsInventoryStatus.LockType = 3;
rsInventoryStatus.Open();
%>
<html>
<head>
	<title>Inventory - Class Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript" src="../js/m003Srh02.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	   

	function initscr() {
		document.frm03s03.ClassSearchPickList.focus();
	}
	
	function CnstrFltr() {	
		var inspSrtBy = document.frm03s03.SortByColumn.value;
		var inspSrtOrd = document.frm03s03.OrderBy.value;
		var stgFilter = "insCurrent_Status = " + document.frm03s03.LookupValueOptions.value;
		var ClassID = document.frm03s03.ClassSearchID.value;
		var ClassType = document.frm03s03.ClassType.value;
		if ((ClassID > 0) && (ClassType != "")) {
			document.frm03s03.action = "m003q03.asp?ClassID="+ClassID+"&ClassType="+ClassType+"&inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;		
			document.frm03s03.submit() ; 		
		} else {
			alert("Please select a class.");
			return ;
		}
	}

	function Toggle() {
		var idx = document.frm03s03.ClassSearchOperand.value;
		switch (idx) {
			// class no 
			case "39":
				openWindow("m003p0103.asp","");
			break;
			default: 
				document.frm03s03.ClassSearchText.value = ""; 
				document.frm03s03.ClassType.value = ""; 				
			break;
	   }
	}

	</script>
</head>
<body onload="initscr()" >
<form name="frm03s03" method="post">
<h3>Inventory - Class Search</h3>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><select name="ClassSearchOperand" tabindex="1" style="width: 150px">
				<option value="39">Inventory Class
		</select></td>
		<td>
			<input type="text" name="ClassSearchText" READONLY>
			<input type="button" name="ClassSearchPickList" value="List" onClick="Toggle();" tabindex="2" class="btnstyle">
		</td>
    </tr>
</table><br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><select name="LookupValueOperand" tabindex="3" style="width: 150px">
				<option value="51">Inventory Status
		</select></td>		
		<td><select name="LookupValueOptions" tabindex="4">
			<%
			while (!rsInventoryStatus.EOF) {			
			%>	
				<option value="<%=rsInventoryStatus.Fields.Item("insEquip_status_id").Value%>" <%=((rsInventoryStatus.Fields.Item("insEquip_status_id").Value==1)?"SELECTED":"")%>><%=rsInventoryStatus.Fields.Item("chvStatusDesc").Value%>
			<%
				rsInventoryStatus.MoveNext();
			}
			%>	
		</select></td>		
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>
			Sort by:
			<select name="SortByColumn" tabindex="5">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value == 1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
			Order 
        	<select name="OrderBy" tabindex="6">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>
	        <input type="button" value="Search" onClick="CnstrFltr();" tabindex="7" class="btnstyle">
			<input type="reset" value="Clear All" onClick="window.location.reload();" tabindex="8" class="btnstyle">
		</td>		
    </tr>
</table>
<input type="hidden" name="ClassSearchID">
<input type="hidden" name="ClassType">
</form>
</body>
</html>
<%
rsCol.Close();
%>
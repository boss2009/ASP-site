<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsInventory__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsInventory__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}

var rsInventory__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsInventory__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsInventory__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsInventory__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsInventory = Server.CreateObject("ADODB.Recordset");
rsInventory.ActiveConnection = MM_cnnASP02_STRING;
rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory_02("+ rsInventory__inspSrtBy.replace(/'/g, "''") + ","+ rsInventory__inspSrtOrd.replace(/'/g, "''") + ",'"+ rsInventory__chvFilter.replace(/'/g, "''") + "',0,0,0)}";
rsInventory.CursorType = 0;
rsInventory.CursorLocation = 2;
rsInventory.LockType = 3;
rsInventory.Open();
var rsInventory_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsInventory_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsInventory_total = rsInventory.RecordCount;

// set the number of rows displayed on this page
if (rsInventory_numRows < 0) {            // if repeat region set to all records
  rsInventory_numRows = rsInventory_total;
} else if (rsInventory_numRows == 0) {    // if no repeat regions
  rsInventory_numRows = 1;
}

// set the first and last displayed record
var rsInventory_first = 1;
var rsInventory_last  = rsInventory_first + rsInventory_numRows - 1;

// if we have the correct record count, check the other stats
if (rsInventory_total != -1) {
  rsInventory_numRows = Math.min(rsInventory_numRows, rsInventory_total);
  rsInventory_first   = Math.min(rsInventory_first, rsInventory_total);
  rsInventory_last    = Math.min(rsInventory_last, rsInventory_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsInventory_total == -1) {

  // count the total records by iterating through the recordset
  for (rsInventory_total=0; !rsInventory.EOF; rsInventory.MoveNext()) {
    rsInventory_total++;
  }

  // reset the cursor to the beginning
  if (rsInventory.CursorType > 0) {
    if (!rsInventory.BOF) rsInventory.MoveFirst();
  } else {
    rsInventory.Requery();
  }

  // set the number of rows displayed on this page
  if (rsInventory_numRows < 0 || rsInventory_numRows > rsInventory_total) {
    rsInventory_numRows = rsInventory_total;
  }

  // set the first and last displayed record
  rsInventory_last  = Math.min(rsInventory_first + rsInventory_numRows - 1, rsInventory_total);
  rsInventory_first = Math.min(rsInventory_first, rsInventory_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsInventory;
var MM_rsCount   = rsInventory_total;
var MM_size      = rsInventory_numRows;
var MM_uniqueCol = "";
    MM_paramName = "";
var MM_offset = 0;
var MM_atTotal = false;
var MM_paramIsDefined = (MM_paramName != "" && String(Request(MM_paramName)) != "undefined");

// *** Move To Record: handle 'index' or 'offset' parameter

if (!MM_paramIsDefined && MM_rsCount != 0) {

  // use index parameter if defined, otherwise use offset parameter
  r = String(Request("index"));
  if (r == "undefined") r = String(Request("offset"));
  if (r && r != "undefined") MM_offset = parseInt(r);

  // if we have a record count, check if we are past the end of the recordset
  if (MM_rsCount != -1) {
    if (MM_offset >= MM_rsCount || MM_offset == -1) {  // past end or move last
      if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount % MM_size);
      } else {
        MM_offset = MM_rsCount - MM_size;
      }
    }
  }

  // move the cursor to the selected record
  for (var i=0; !MM_rs.EOF && (i < MM_offset || MM_offset == -1); i++) {
    MM_rs.MoveNext();
  }
  if (MM_rs.EOF) MM_offset = i;  // set MM_offset to the last possible record
}

// *** Move To Record: if we dont know the record count, check the display range

if (MM_rsCount == -1) {

  // walk to the end of the display range for this page
  for (var i=MM_offset; !MM_rs.EOF && (MM_size < 0 || i < MM_offset + MM_size); i++) {
    MM_rs.MoveNext();
  }

  // if we walked off the end of the recordset, set MM_rsCount and MM_size
  if (MM_rs.EOF) {
    MM_rsCount = i;
    if (MM_size < 0 || MM_size > MM_rsCount) MM_size = MM_rsCount;
  }

  // if we walked off the end, set the offset based on page size
  if (MM_rs.EOF && !MM_paramIsDefined) {
    if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
      MM_offset = MM_rsCount - (MM_rsCount % MM_size);
    } else {
      MM_offset = MM_rsCount - MM_size;
    }
  }

  // reset the cursor to the beginning
  if (MM_rs.CursorType > 0) {
    if (!MM_rs.BOF) MM_rs.MoveFirst();
  } else {
    MM_rs.Requery();
  }

  // move the cursor to the selected record
  for (var i=0; !MM_rs.EOF && i < MM_offset; i++) {
    MM_rs.MoveNext();
  }
}
// *** Move To Record: update recordset stats

// set the first and last displayed record
rsInventory_first = MM_offset + 1;
rsInventory_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsInventory_first = Math.min(rsInventory_first, MM_rsCount);
  rsInventory_last  = Math.min(rsInventory_last, MM_rsCount);
}

// set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount != -1 && MM_offset + MM_size >= MM_rsCount);
// *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

// create the list of parameters which should not be maintained
var MM_removeList = "&index=";
if (MM_paramName != "") MM_removeList += "&" + MM_paramName.toLowerCase() + "=";
var MM_keepURL="",MM_keepForm="",MM_keepBoth="",MM_keepNone="";

// add the URL parameters to the MM_keepURL string
for (var items=new Enumerator(Request.QueryString); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepURL += "&" + items.item() + "=" + Server.URLencode(Request.QueryString(items.item()));
  }
}

// add the Form variables to the MM_keepForm string
for (var items=new Enumerator(Request.Form); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepForm += "&" + items.item() + "=" + Server.URLencode(Request.Form(items.item()));
  }
}

// create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL + MM_keepForm;
if (MM_keepBoth.length > 0) MM_keepBoth = MM_keepBoth.substring(1);
if (MM_keepURL.length > 0)  MM_keepURL = MM_keepURL.substring(1);
if (MM_keepForm.length > 0) MM_keepForm = MM_keepForm.substring(1);
// *** Move To Record: set the strings for the first, last, next, and previous links

var MM_moveFirst="",MM_moveLast="",MM_moveNext="",MM_movePrev="";
var MM_keepMove = MM_keepBoth;  // keep both Form and URL parameters for moves
var MM_moveParam = "index";

// if the page has a repeated region, remove 'offset' from the maintained parameters
if (MM_size > 1) {
  MM_moveParam = "offset";
  if (MM_keepMove.length > 0) {
    params = MM_keepMove.split("&");
    MM_keepMove = "";
    for (var i=0; i < params.length; i++) {
      var nextItem = params[i].substring(0,params[i].indexOf("="));
      if (nextItem.toLowerCase() != MM_moveParam) {
        MM_keepMove += "&" + params[i];
      }
    }
    if (MM_keepMove.length > 0) MM_keepMove = MM_keepMove.substring(1);
  }
}

// set the strings for the move to links
if (MM_keepMove.length > 0) MM_keepMove += "&";
var urlStr = Request.ServerVariables("URL") + "?" + MM_keepMove + MM_moveParam + "=";
MM_moveFirst = urlStr + "0";
MM_moveLast  = urlStr + "-1";
MM_moveNext  = urlStr + (MM_offset + MM_size);
MM_movePrev  = urlStr + Math.max(MM_offset - MM_size,0);
%>
<html>
<head>
	<title>Inventory - Browse</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>		
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	
	function JumpRecord(){
		if (document.frmq01.JumpToRecord.value=="") {
			alert("Enter Record Number.");
		}
		if (!IsID(document.frmq01.JumpToRecord.value)) {
			alert("Invalid Record Number.");
		} else {
			window.location.href="..<%Response.Write(Request.ServerVariables("URL") + "?" + MM_keepMove + MM_moveParam + "=")%>"+String(document.frmq01.JumpToRecord.value-1);
		}
	}			
	</Script>
</head>
<body>
<form name="frmq01">
<h3>Inventory - Browse</h3>
<table cellspacing="1">
    <tr> 
		<td align="left" width="900">Displaying Records <b><%=(rsInventory_first)%></b> to <b><%=(rsInventory_last)%></b> of <b><%=(rsInventory_total)%></b></td>
		<td nowrap><a href="javascript: openWindow('m003a0101.asp','w003A01');">Add Inventory</a></td>
	</tr>
</table>
<div class="BrowsePanel" style="width: 100%; height: 295px"> 
<table cellpadding="2" cellspacing="1">
	<tr> 
        <th nowrap class="headrow" align="left" width="280">Inventory Name</th>
        <th nowrap class="headrow" align="left">Inventory ID</th>
        <th nowrap class="headrow" align="left">Current Status</th>		
        <th nowrap class="headrow" align="left">Model Number</th>
        <th nowrap class="headrow" align="left">Serial Number</th>
        <th nowrap class="headrow" align="left">PR Number</th>
        <th nowrap class="headrow" align="left" width="200">Vendor</th>
        <th nowrap class="headrow" align="left">Current User</th>
        <th nowrap class="headrow" align="left">Inventory Cost</th>
        <th nowrap class="headrow" align="left">Sold Cost</th>
	</tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsInventory.EOF)) { 
%>
	<tr> 
        <td align="left" valign="top"><a href="javascript: openWindow('m003FS3.asp?intEquip_Set_id=<%=(rsInventory.Fields.Item("intEquip_Set_id").Value)%>&intBar_Code_no=<%=rsInventory.Fields.Item("intBar_Code_no").Value%>','w003E01');"><%=Truncate(rsInventory.Fields.Item("chvInventory_Name").Value,40)%></a>&nbsp;</td>
        <td align="center" valign="top" nowrap><%=ZeroPadFormat(rsInventory.Fields.Item("intBar_Code_no").Value,8)%>&nbsp;</td>
        <td align="left" valign="top" nowrap><%=(rsInventory.Fields.Item("chvEqp_Status").Value)%>&nbsp;</td>		
        <td align="left" valign="top" nowrap><%=rsInventory.Fields.Item("chvModel_Number").Value%>&nbsp;</td>
        <td align="left" valign="top" nowrap><%=rsInventory.Fields.Item("chvSerial_Number").Value%>&nbsp;</td>
        <td align="center" valign="top" nowrap><%=ZeroPadFormat(rsInventory.Fields.Item("intRequisition_no").Value,8)%>&nbsp;</td>
        <td align="left" valign="top"><%=rsInventory.Fields.Item("chvVendor_Name").Value%>&nbsp;</td>
        <td align="left" valign="top" nowrap><%=((rsInventory.Fields.Item("insUser_Type_id").Value==4)?rsInventory.Fields.Item("chvInstitUsr_Nm").Value:rsInventory.Fields.Item("chvUsr_Lst_Name").Value+", "+rsInventory.Fields.Item("chvUsr_Fst_Name").Value)%>&nbsp;</td>
        <td align="right" valign="top" nowrap><%=FormatCurrency(rsInventory.Fields.Item("fltPurchase_Cost").Value)%>&nbsp;</td>
        <td align="right" valign="top" nowrap><%=FormatCurrency(rsInventory.Fields.Item("fltEqp_Sold_Price").Value)%>&nbsp;</td>
	</tr>
<%
	Repeat1__index++;
	rsInventory.MoveNext();
}
%>
</table>
</div>
</form>
</body>
</html>
<%
rsInventory.Close();
%>
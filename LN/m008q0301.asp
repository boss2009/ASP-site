<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var gst = 0;
var pst = 0;
var shipping = 0;

var rsGST = Server.CreateObject("ADODB.Recordset");
rsGST.ActiveConnection = MM_cnnASP02_STRING;
rsGST.Source = "{call dbo.cp_charge_rate(1,'',0,0.0,1,'Q',0)}";
rsGST.CursorType = 0;
rsGST.CursorLocation = 2;
rsGST.LockType = 3;
rsGST.Open();
if (!rsGST.EOF) gst = Number(rsGST.Fields.Item("fltPercentage").Value);
rsGST.Close();

var rsPST = Server.CreateObject("ADODB.Recordset");
rsPST.ActiveConnection = MM_cnnASP02_STRING;
rsPST.Source = "{call dbo.cp_charge_rate(2,'',0,0.0,1,'Q',0)}";
rsPST.CursorType = 0;
rsPST.CursorLocation = 2;
rsPST.LockType = 3;
rsPST.Open();
if (!rsPST.EOF) pst = Number(rsPST.Fields.Item("fltPercentage").Value);
rsPST.Close();

var rsShipping = Server.CreateObject("ADODB.Recordset");
rsShipping.ActiveConnection = MM_cnnASP02_STRING;
rsShipping.Source = "{call dbo.cp_charge_rate(3,'',0,0.0,1,'Q',0)}";
rsShipping.CursorType = 0;
rsShipping.CursorLocation = 2;
rsShipping.LockType = 3;
rsShipping.Open();
if (!rsShipping.EOF) shipping = Number(rsShipping.Fields.Item("fltPercentage").Value);
rsShipping.Close();

var rsInventoryLoaned = Server.CreateObject("ADODB.Recordset");
rsInventoryLoaned.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryLoaned.Source = "{call dbo.cp_eqp_loaned(0,"+Request.QueryString("intLoan_Req_id")+",0,'',0,0,'','',0,'Q',0)}";
rsInventoryLoaned.CursorType = 0;
rsInventoryLoaned.CursorLocation = 2;
rsInventoryLoaned.LockType = 3;
rsInventoryLoaned.Open();
var rsInventoryLoaned_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsInventoryLoaned_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsInventoryLoaned_total = rsInventoryLoaned.RecordCount;

// set the number of rows displayed on this page
if (rsInventoryLoaned_numRows < 0) {            // if repeat region set to all records
  rsInventoryLoaned_numRows = rsInventoryLoaned_total;
} else if (rsInventoryLoaned_numRows == 0) {    // if no repeat regions
  rsInventoryLoaned_numRows = 1;
}

// set the first and last displayed record
var rsInventoryLoaned_first = 1;
var rsInventoryLoaned_last  = rsInventoryLoaned_first + rsInventoryLoaned_numRows - 1;

// if we have the correct record count, check the other stats
if (rsInventoryLoaned_total != -1) {
  rsInventoryLoaned_numRows = Math.min(rsInventoryLoaned_numRows, rsInventoryLoaned_total);
  rsInventoryLoaned_first   = Math.min(rsInventoryLoaned_first, rsInventoryLoaned_total);
  rsInventoryLoaned_last    = Math.min(rsInventoryLoaned_last, rsInventoryLoaned_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsInventoryLoaned_total == -1) {

  // count the total records by iterating through the recordset
  for (rsInventoryLoaned_total=0; !rsInventoryLoaned.EOF; rsInventoryLoaned.MoveNext()) {
    rsInventoryLoaned_total++;
  }

  // reset the cursor to the beginning
  if (rsInventoryLoaned.CursorType > 0) {
    if (!rsInventoryLoaned.BOF) rsInventoryLoaned.MoveFirst();
  } else {
    rsInventoryLoaned.Requery();
  }

  // set the number of rows displayed on this page
  if (rsInventoryLoaned_numRows < 0 || rsInventoryLoaned_numRows > rsInventoryLoaned_total) {
    rsInventoryLoaned_numRows = rsInventoryLoaned_total;
  }

  // set the first and last displayed record
  rsInventoryLoaned_last  = Math.min(rsInventoryLoaned_first + rsInventoryLoaned_numRows - 1, rsInventoryLoaned_total);
  rsInventoryLoaned_first = Math.min(rsInventoryLoaned_first, rsInventoryLoaned_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsInventoryLoaned;
var MM_rsCount   = rsInventoryLoaned_total;
var MM_size      = rsInventoryLoaned_numRows;
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
rsInventoryLoaned_first = MM_offset + 1;
rsInventoryLoaned_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsInventoryLoaned_first = Math.min(rsInventoryLoaned_first, MM_rsCount);
  rsInventoryLoaned_last  = Math.min(rsInventoryLoaned_last, MM_rsCount);
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
	<title>Equipment Loaned</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>
</head>
<body>
<h5>Equipment Loaned</h5>
<table cellspacing="1">
    <tr> 
		<td nowrap width="450"><a href="javascript: openWindow('m008a0301.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>','w008A02');">Add Equipment Loaned</a></td>	
    	<td nowrap align="left">Displaying <b><%=(rsInventoryLoaned_total)%></b> Records.</td>
	</tr>
</table>
<hr>
<div class="BrowsePanel" style="width: 564px; height: 219px"> 
  <table cellpadding="2" cellspacing="1">
    <tr> 
      <th nowrap class="headrow" align="left">Inventory Name</th>
      <th nowrap class="headrow" align="left">Inventory ID</th>
      <th nowrap class="headrow" align="left">Date Returned</th>
      <th nowrap class="headrow" align="left">Status</th>
      <th nowrap class="headrow" align="left">Vendor</th>
      <th nowrap class="headrow" align="left">Model Number</th>
      <th nowrap class="headrow" align="left">Serial Number</th>
      <th nowrap class="headrow" align="left">PR Number</th>
      <th nowrap class="headrow" align="left">Equipment Cost</th>
      <th nowrap class="headrow" align="left">Date Processed</th>
      <th nowrap class="headrow" align="left">Returned By</th>
      <th nowrap class="headrow" align="left">Return Status</th>
      <th nowrap class="headrow" align="left">Comments</th>
    </tr>
<% 
var tax = 0;
var total_shipping = 0;
var total_cost = 0;
var total_loan = 0;
while ((!rsInventoryLoaned.EOF)) { 		
	if (!(rsInventoryLoaned.Fields.Item("insEquip_Class_id").Value==null)) {		
		var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
		rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
		rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW(" + rsInventoryLoaned.Fields.Item("insEquip_Class_id").Value + ",'C',1)}";	
		rsConcreteClass.CursorType = 0;
		rsConcreteClass.CursorLocation = 2;
		rsConcreteClass.LockType = 3;
		rsConcreteClass.Open();	
		switch (String(rsConcreteClass.Fields.Item("chvSbjTotax").Value)) {
			//pst
			case "1":
				tax = tax + (rsInventoryLoaned.Fields.Item("fltList_Unit_Cost").Value * (pst/100));
			break;
			//gst
			case "2":
				tax = tax + (rsInventoryLoaned.Fields.Item("fltList_Unit_Cost").Value * (gst/100));		
			break;
			//both
			case "3":
				tax = tax + (rsInventoryLoaned.Fields.Item("fltList_Unit_Cost").Value * ((gst+pst)/100));		
			break;
		}
	}
%>
    <tr> 
<!-- + Nov.03.2005
-->
      <td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvInventory_Name").Value)%>&nbsp;</td>

      <td nowrap align="center"><%=ZeroPadFormat(rsInventoryLoaned.Fields.Item("intBar_Code_no").Value,8)%>&nbsp;</td>
      <td nowrap align="center"><%=FilterDate(rsInventoryLoaned.Fields.Item("dtsDate_Returned").Value)%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvInventory_Status").Value)%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvVendor_Name").Value)%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvModel_Number").Value)%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvSerial_Number").Value)%>&nbsp;</td>
      <td nowrap align="center"><%=ZeroPadFormat(rsInventoryLoaned.Fields.Item("intRequisition_no").Value,8)%>&nbsp;</td>
      <td nowrap align="right"><%=FormatCurrency(rsInventoryLoaned.Fields.Item("fltList_Unit_Cost").Value)%>&nbsp;</td>
      <td nowrap align="center"><%=FilterDate(rsInventoryLoaned.Fields.Item("dtsDate_Shipped").Value)%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvReturned_by").Value)%>&nbsp;</td>
      <td nowrap align="center"><%=(rsInventoryLoaned.Fields.Item("chvRtn_Complete").Value)%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvComments").Value)%>&nbsp;</td>
    </tr>
<%
if (rsInventoryLoaned.Fields.Item("dtsDate_Returned").Value==null) total_loan += rsInventoryLoaned.Fields.Item("fltList_Unit_Cost").Value;
	total_cost += rsInventoryLoaned.Fields.Item("fltList_Unit_Cost").Value;
	rsInventoryLoaned.MoveNext();
}

total_shipping = total_cost * (shipping/100);
%>
  </table>
</div>
<div style="position: absolute; top: 310px">
<table cellpadding="0" cellspacing="1">
	<tr>
		<td width="350"><b>Total Loan Cost with taxes/shipping:</b></td>
		<td align="right"><b><%=FormatCurrency(total_cost+ tax + total_shipping)%></b></td>
	</tr>
	<tr>
		<td width="350"><b>Total Cost of Equipment Still On Loan:</b></td>
		<td align="right"><b><%=FormatCurrency(total_loan)%></b></td>
	</tr>
</table>
</div>
</body>
</html>
<%
rsInventoryLoaned.Close();
%>
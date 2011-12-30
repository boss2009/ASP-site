<!--------------------------------------------------------------------------
* File Name: m014q0301.asp
* Title: Equipment Received
* Main SP: cp_purchase_requisition_received
* Description: This page lists all the equipment received.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsInventoryReceived = Server.CreateObject("ADODB.Recordset");
rsInventoryReceived.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryReceived.Source = "{call dbo.cp_Purchase_Requisition_Received("+Request.QueryString("insPurchase_Req_id")+",0,0,'',0,'',0,'Q',0)}";
rsInventoryReceived.CursorType = 0;
rsInventoryReceived.CursorLocation = 2;
rsInventoryReceived.LockType = 3;
rsInventoryReceived.Open();

var rsRequisition = Server.CreateObject("ADODB.Recordset");
rsRequisition.ActiveConnection = MM_cnnASP02_STRING;
rsRequisition.Source = "{call dbo.cp_Get_Purchase_Requisition(0,0,'',1,"+ Request.QueryString("insPurchase_Req_id")+ ",0)}";
rsRequisition.CursorType = 0;
rsRequisition.CursorLocation = 2;
rsRequisition.LockType = 3;
rsRequisition.Open();

var rsInventoryReceived_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsInventoryReceived_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsInventoryReceived_total = rsInventoryReceived.RecordCount;

// set the number of rows displayed on this page
if (rsInventoryReceived_numRows < 0) {            // if repeat region set to all records
  rsInventoryReceived_numRows = rsInventoryReceived_total;
} else if (rsInventoryReceived_numRows == 0) {    // if no repeat regions
  rsInventoryReceived_numRows = 1;
}

// set the first and last displayed record
var rsInventoryReceived_first = 1;
var rsInventoryReceived_last  = rsInventoryReceived_first + rsInventoryReceived_numRows - 1;

// if we have the correct record count, check the other stats
if (rsInventoryReceived_total != -1) {
  rsInventoryReceived_numRows = Math.min(rsInventoryReceived_numRows, rsInventoryReceived_total);
  rsInventoryReceived_first   = Math.min(rsInventoryReceived_first, rsInventoryReceived_total);
  rsInventoryReceived_last    = Math.min(rsInventoryReceived_last, rsInventoryReceived_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsInventoryReceived_total == -1) {

  // count the total records by iterating through the recordset
  for (rsInventoryReceived_total=0; !rsInventoryReceived.EOF; rsInventoryReceived.MoveNext()) {
    rsInventoryReceived_total++;
  }

  // reset the cursor to the beginning
  if (rsInventoryReceived.CursorType > 0) {
    if (!rsInventoryReceived.BOF) rsInventoryReceived.MoveFirst();
  } else {
    rsInventoryReceived.Requery();
  }

  // set the number of rows displayed on this page
  if (rsInventoryReceived_numRows < 0 || rsInventoryReceived_numRows > rsInventoryReceived_total) {
    rsInventoryReceived_numRows = rsInventoryReceived_total;
  }

  // set the first and last displayed record
  rsInventoryReceived_last  = Math.min(rsInventoryReceived_first + rsInventoryReceived_numRows - 1, rsInventoryReceived_total);
  rsInventoryReceived_first = Math.min(rsInventoryReceived_first, rsInventoryReceived_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsInventoryReceived;
var MM_rsCount   = rsInventoryReceived_total;
var MM_size      = rsInventoryReceived_numRows;
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
rsInventoryReceived_first = MM_offset + 1;
rsInventoryReceived_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsInventoryReceived_first = Math.min(rsInventoryReceived_first, MM_rsCount);
  rsInventoryReceived_last  = Math.min(rsInventoryReceived_last, MM_rsCount);
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
	<title>Equipment Received</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>
</head>
<body>
<h5>Equipment Received</h5>
<%
if ((rsRequisition.Fields.Item("insPurchase_sts_id").Value!=6) && (rsRequisition.Fields.Item("insPurchase_sts_id").Value!=7)) {
%>
<i>Change status to Received Complete or Received Incomplete on General Information Page before receiving inventory.</i>
<%
} else {
%>
<table cellspacing="1">
	<tr> 
		<td colspan="4" align="left">Displaying <b><%=(rsInventoryReceived_total)%></b> Records.</td>
	</tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr> 
		<th class="headrow" align="left" valign="top" nowrap>Class/Bundle</th>
		<th class="headrow" align="left" valign="top">Description</th>
		<th class="headrow" align="left" valign="top">Quantity Ordered</th>
		<th class="headrow" align="left" valign="top">Quantity Received</th>	  
		<th class="headrow" align="left" valign="top">Date Received</th>	  
		<th class="headrow" align="left" valign="top">List Unit Price</th>
		<th class="headrow" align="left" valign="top" nowrap>Sub Total</th>
    </tr>
<% 
var total_cost = 0;
while (!rsInventoryReceived.EOF) { 
%>
    <tr>
		<td valign="top" nowrap><a href="m014e0301.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>&insRqst_received_id=<%=(rsInventoryReceived.Fields.Item("insRqst_received_id").Value)%>"><%=(rsInventoryReceived.Fields.Item("chvClass_name").Value)%></a></td>
		<td valign="top" align="left"><%=(rsInventoryReceived.Fields.Item("chvNotes").Value)%>&nbsp;</td>
		<td valign="top" align="center" nowrap><%=(rsInventoryReceived.Fields.Item("intQuantity_Ordered").Value)%>&nbsp;</td>
		<td valign="top" align="center" nowrap><%=(rsInventoryReceived.Fields.Item("intQuantity_Received").Value)%>&nbsp;</td>
		<td valign="top" align="center" nowrap><%=FilterDate(rsInventoryReceived.Fields.Item("dtsReceived").Value)%>&nbsp;</td>
		<td valign="top" align="right" nowrap><%=FormatCurrency(rsInventoryReceived.Fields.Item("fltList_unit_cost").Value)%>&nbsp;</td>
		<td valign="top" align="right" nowrap><%=FormatCurrency(rsInventoryReceived.Fields.Item("fltTotal_cost").Value)%>&nbsp;</td>
    </tr>    
<%
	total_cost += rsInventoryReceived.Fields.Item("fltTotal_cost").Value;
	rsInventoryReceived.MoveNext();
}
%>
    <tr> 
		<td colspan="5"></td>
		<td nowrap align="right"><b>Total:</b></td>
		<td nowrap align="right"><b><%=FormatCurrency(total_cost)%></b></td>
    </tr>
</table>
<%
}
%>
</body>
</html>
<%
rsInventoryReceived.Close();
%>
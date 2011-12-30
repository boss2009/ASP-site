<!--------------------------------------------------------------------------
* File Name: m014r0102excel.asp
* Title: Delivery Performance
* Main SP: cp_PR_Pfrm_Rpt
* Description: Delivery Performance Report in excel format.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsPerformance = Server.CreateObject("ADODB.Recordset");
rsPerformance.ActiveConnection = MM_cnnASP02_STRING;
rsPerformance.Source = "{call dbo.cp_PR_Pfrm_Rpt("+Request.QueryString("Vendor")+",0,'"+Request.QueryString("StartDate")+"','"+Request.QueryString("EndDate")+"',0,0,0)}";
rsPerformance.CursorType = 0;
rsPerformance.CursorLocation = 2;
rsPerformance.LockType = 3;
rsPerformance.Open();
var rsPerformance_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsPerformance_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsPerformance_total = rsPerformance.RecordCount;

// set the number of rows displayed on this page
if (rsPerformance_numRows < 0) {            // if repeat region set to all records
  rsPerformance_numRows = rsPerformance_total;
} else if (rsPerformance_numRows == 0) {    // if no repeat regions
  rsPerformance_numRows = 1;
}

// set the first and last displayed record
var rsPerformance_first = 1;
var rsPerformance_last  = rsPerformance_first + rsPerformance_numRows - 1;

// if we have the correct record count, check the other stats
if (rsPerformance_total != -1) {
  rsPerformance_numRows = Math.min(rsPerformance_numRows, rsPerformance_total);
  rsPerformance_first   = Math.min(rsPerformance_first, rsPerformance_total);
  rsPerformance_last    = Math.min(rsPerformance_last, rsPerformance_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsPerformance_total == -1) {

  // count the total records by iterating through the recordset
  for (rsPerformance_total=0; !rsPerformance.EOF; rsPerformance.MoveNext()) {
    rsPerformance_total++;
  }

  // reset the cursor to the beginning
  if (rsPerformance.CursorType > 0) {
    if (!rsPerformance.BOF) rsPerformance.MoveFirst();
  } else {
    rsPerformance.Requery();
  }

  // set the number of rows displayed on this page
  if (rsPerformance_numRows < 0 || rsPerformance_numRows > rsPerformance_total) {
    rsPerformance_numRows = rsPerformance_total;
  }

  // set the first and last displayed record
  rsPerformance_last  = Math.min(rsPerformance_first + rsPerformance_numRows - 1, rsPerformance_total);
  rsPerformance_first = Math.min(rsPerformance_first, rsPerformance_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsPerformance;
var MM_rsCount   = rsPerformance_total;
var MM_size      = rsPerformance_numRows;
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
rsPerformance_first = MM_offset + 1;
rsPerformance_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsPerformance_first = Math.min(rsPerformance_first, MM_rsCount);
  rsPerformance_last  = Math.min(rsPerformance_last, MM_rsCount);
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
for (var items=new Enumerator(Request.QueryString); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepForm += "&" + items.item() + "=" + Server.URLencode(Request.QueryString(items.item()));
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
	<title>Delivery Performance-Vendor: <%=Request.QueryString("VendorName")%></title>
</head>
<body>
<table>
	<tr>
		<td colspan="7">Delivery Performance - <%=Request.QueryString("VendorName")%></td>
	</tr>
	<tr>
		<td colspan="7"><%=Request.QueryString("StartDate")%> to <%=Request.QueryString("EndDate")%>: <%=(rsPerformance_total)%> Purchase Requisitions</td>
	</tr>
    <tr> 
		<th>Inventory Class</th>	
		<th>PR Number</th>
		<th>Date Ordered</th>
		<th>Date Received</th>
		<th>Days Taken</th>	  
		<th>Backorder Received Date</th>	  
		<th>B/O Equip./Days Taken</th>
    </tr>
<% 
var total1 = 0;
var total2 = 0;	
while (!rsPerformance.EOF) { 
%>
    <tr> 
		<td><%=rsPerformance.Fields.Item("chvName").Value%></td>
		<td><%=(rsPerformance.Fields.Item("insPurchase_req_id").Value)%>&nbsp;</td>
		<td><%=(rsPerformance.Fields.Item("dtsDate_Ordered").Value)%>&nbsp;</td>
		<td><%=(rsPerformance.Fields.Item("dtsReceived").Value)%>&nbsp;</td>
		<td><%=(rsPerformance.Fields.Item("intDays_taken").Value)%>&nbsp;</td>
		<td><%=(rsPerformance.Fields.Item("dtsBack_Ord_rx").Value)%>&nbsp;</td>
		<td><%=(rsPerformance.Fields.Item("intBODays_taken").Value)%>&nbsp;</td>	  
    </tr>
<%
	total1=total1+rsPerformance.Fields.Item("intDays_taken").Value;
	total2=total2+rsPerformance.Fields.Item("intDays_taken").Value+rsPerformance.Fields.Item("intBODays_taken").Value;
	rsPerformance.MoveNext();
}
if (rsPerformance_total > 0) {
%>
  	<tr>
		<td></td>
	</tr>
  	<tr>
		<td colspan="7">Average # of Days Per Product = <%=(total1/rsPerformance_total)%>(not including B/O Equipment)</td>
	</tr>
	<tr>
		<td colspan="7">Average # of Days Per Product = <%=(total2/rsPerformance_total)%>(including B/O Equipment)</td>
	</tr>
<%
}
%>
</table>
</body>
</html>
<%
rsPerformance.Close();
%>
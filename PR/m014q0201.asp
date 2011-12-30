<!--------------------------------------------------------------------------
* File Name: m014q0201.asp
* Title: Equipment Requested
* Main SP: cp_purchase_requisition_requested
* Description: This page lists equipment requested of a purchase requisition.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.QueryString("Action"))=="Remove") {
	var cmdRemoveClass = Server.CreateObject("ADODB.Command");
	cmdRemoveClass.ActiveConnection = MM_cnnASP02_STRING;
	cmdRemoveClass.CommandText = "dbo.cp_Delete_Rqst_Requested";
	cmdRemoveClass.CommandType = 4;
	cmdRemoveClass.CommandTimeout = 0;
	cmdRemoveClass.Prepared = true;
	cmdRemoveClass.Parameters.Append(cmdRemoveClass.CreateParameter("RETURN_VALUE", 3, 4));
	cmdRemoveClass.Parameters.Append(cmdRemoveClass.CreateParameter("@insPurchase_req_id", 2, 1,1,Request.QueryString("insPurchase_req_id")));
	cmdRemoveClass.Parameters.Append(cmdRemoveClass.CreateParameter("@insEquip_class_id", 2, 1,1,Request.QueryString("insClass_bundle_id")));
	cmdRemoveClass.Parameters.Append(cmdRemoveClass.CreateParameter("@bitIs_class", 2, 1,1,1));
	cmdRemoveClass.Parameters.Append(cmdRemoveClass.CreateParameter("@insRtnFlag", 2, 2));
	cmdRemoveClass.Execute();
	if (cmdRemoveClass.Parameters.Item("@insRtnFlag").Value!=0) Response.Redirect("RemoveClassFailed.asp?Page=m014q0201.asp&insPurchase_req_id="+Request.QueryString("insPurchase_req_id"));
}

var rsInventoryRequested = Server.CreateObject("ADODB.Recordset");
rsInventoryRequested.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryRequested.Source = "{call dbo.cp_Purchase_Requisition_Requested2(0,"+Request.QueryString("insPurchase_Req_id")+",0,0,0,'',0.0,'',0,0,0,0,0,'Q',0)}";
rsInventoryRequested.CursorType = 0;
rsInventoryRequested.CursorLocation = 2;
rsInventoryRequested.LockType = 3;
rsInventoryRequested.Open();
var rsInventoryRequested_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsInventoryRequested_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsInventoryRequested_total = rsInventoryRequested.RecordCount;

// set the number of rows displayed on this page
if (rsInventoryRequested_numRows < 0) {            // if repeat region set to all records
  rsInventoryRequested_numRows = rsInventoryRequested_total;
} else if (rsInventoryRequested_numRows == 0) {    // if no repeat regions
  rsInventoryRequested_numRows = 1;
}

// set the first and last displayed record
var rsInventoryRequested_first = 1;
var rsInventoryRequested_last  = rsInventoryRequested_first + rsInventoryRequested_numRows - 1;

// if we have the correct record count, check the other stats
if (rsInventoryRequested_total != -1) {
  rsInventoryRequested_numRows = Math.min(rsInventoryRequested_numRows, rsInventoryRequested_total);
  rsInventoryRequested_first   = Math.min(rsInventoryRequested_first, rsInventoryRequested_total);
  rsInventoryRequested_last    = Math.min(rsInventoryRequested_last, rsInventoryRequested_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsInventoryRequested_total == -1) {

  // count the total records by iterating through the recordset
  for (rsInventoryRequested_total=0; !rsInventoryRequested.EOF; rsInventoryRequested.MoveNext()) {
    rsInventoryRequested_total++;
  }

  // reset the cursor to the beginning
  if (rsInventoryRequested.CursorType > 0) {
    if (!rsInventoryRequested.BOF) rsInventoryRequested.MoveFirst();
  } else {
    rsInventoryRequested.Requery();
  }

  // set the number of rows displayed on this page
  if (rsInventoryRequested_numRows < 0 || rsInventoryRequested_numRows > rsInventoryRequested_total) {
    rsInventoryRequested_numRows = rsInventoryRequested_total;
  }

  // set the first and last displayed record
  rsInventoryRequested_last  = Math.min(rsInventoryRequested_first + rsInventoryRequested_numRows - 1, rsInventoryRequested_total);
  rsInventoryRequested_first = Math.min(rsInventoryRequested_first, rsInventoryRequested_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsInventoryRequested;
var MM_rsCount   = rsInventoryRequested_total;
var MM_size      = rsInventoryRequested_numRows;
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
rsInventoryRequested_first = MM_offset + 1;
rsInventoryRequested_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsInventoryRequested_first = Math.min(rsInventoryRequested_first, MM_rsCount);
  rsInventoryRequested_last  = Math.min(rsInventoryRequested_last, MM_rsCount);
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
	<title>Equipment Requested</title>
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
<h5>Equipment Requested</h5>
<table cellspacing="1">
	<tr> 
		<td colspan="4" align="left">Displaying <b><%=(rsInventoryRequested_total)%></b> Records.</td>
	</tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr> 
		<th nowrap class="headrow" align="left">Class</th>
		<th nowrap class="headrow" align="left">Description</th>
		<th nowrap class="headrow" align="left">Quantity</th>
		<th nowrap class="headrow" align="left">List Unit Price</th>
		<th nowrap class="headrow" align="left">Sub Total</th>
		<th nowrap class="headrow" align="left">User Name</th>
		<th nowrap class="headrow" align="left">ETA</th>
		<th nowrap class="headrow" align="left">&nbsp;</th>		
	</tr>
<% 
var total_cost = 0;
while (!rsInventoryRequested.EOF) { 
%>
    <tr> 
<% 
	if ((rsInventoryRequested.Fields.Item("chvClass_Bundle_Name").Value==null) || (Trim(rsInventoryRequested.Fields.Item("chvClass_Bundle_Name").Value)=="")) { %>
		<td nowrap><%=(rsInventoryRequested.Fields.Item("insClass_bundle_id").Value)%>&nbsp;</td>
<% 
	} else { 
%>
		<td nowrap><a href="m014e0201.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>&insRqst_requested_id=<%=(rsInventoryRequested.Fields.Item("insRqst_requested_id").Value)%>"><%=(rsInventoryRequested.Fields.Item("chvClass_Bundle_Name").Value)%></a>&nbsp;</td>
<% 
	} 
%>
		<td nowrap><%=rsInventoryRequested.Fields.Item("chvDescription").Value%>&nbsp;</td>
		<td nowrap align="center"><%=(rsInventoryRequested.Fields.Item("insPR_request_Qty_Ordered").Value)%>&nbsp;</td>
		<td nowrap align="right"><%=FormatCurrency(rsInventoryRequested.Fields.Item("fltPR_request_List_Unit_Cost").Value)%>&nbsp;</td>
		<td nowrap align="right"><%=FormatCurrency(rsInventoryRequested.Fields.Item("fltTotal_Cost").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventoryRequested.Fields.Item("chvClient").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=FilterDate(rsInventoryRequested.Fields.Item("dtsEst_delivery_date").Value)%>&nbsp;</td>
<% 
	if ((Session("MM_UserAuthorization") >= 6) && (rsInventoryRequested.Fields.Item("bitFlag_Del").Value=="1")) { 
%>
		<td nowrap><a href="m014q0201.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>&Action=Remove&insClass_bundle_id=<%=(rsInventoryRequested.Fields.Item("insClass_bundle_id").Value)%>"><img src="../i/remove.gif" ALT="Remove <%=(rsInventoryRequested.Fields.Item("chvClass_Bundle_Name").Value)%>"></a></td>
<%
	} 
%>	
	</tr>    
<%
	total_cost += rsInventoryRequested.Fields.Item("fltTotal_Cost").Value;
	rsInventoryRequested.MoveNext();
}
%>
    <tr> 
		<td colspan="3"></td>
		<td nowrap align="right"><b>Total:</b></td>
		<td nowrap align="right"><b><%=FormatCurrency(total_cost)%></b></td>
		<td colspan="2"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><a href="javascript: openWindow('m014a0201.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','w014A02');">Add Inventory Request</a></td>
    </tr>
</table>
</body>
</html>
<%
rsInventoryRequested.Close();
%>
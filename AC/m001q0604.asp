<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsWaiver = Server.CreateObject("ADODB.Recordset");
rsWaiver.ActiveConnection = MM_cnnASP02_STRING;
rsWaiver.Source = "{call dbo.cp_Get_waiver(" + Request.QueryString("intAdult_id") +",0,0)}";
rsWaiver.CursorType = 0;
rsWaiver.CursorLocation = 2;
rsWaiver.LockType = 3;
rsWaiver.Open();
var rsWaiver_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsWaiver_numRows += Repeat1__numRows;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsWaiver_total = rsWaiver.RecordCount;

// set the number of rows displayed on this page
if (rsWaiver_numRows < 0) {            // if repeat region set to all records
  rsWaiver_numRows = rsWaiver_total;
} else if (rsWaiver_numRows == 0) {    // if no repeat regions
  rsWaiver_numRows = 1;
}

// set the first and last displayed record
var rsWaiver_first = 1;
var rsWaiver_last  = rsWaiver_first + rsWaiver_numRows - 1;

// if we have the correct record count, check the other stats
if (rsWaiver_total != -1) {
  rsWaiver_numRows = Math.min(rsWaiver_numRows, rsWaiver_total);
  rsWaiver_first   = Math.min(rsWaiver_first, rsWaiver_total);
  rsWaiver_last    = Math.min(rsWaiver_last, rsWaiver_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsWaiver_total == -1) {

  // count the total records by iterating through the recordset
  for (rsWaiver_total=0; !rsWaiver.EOF; rsWaiver.MoveNext()) {
    rsWaiver_total++;
  }

  // reset the cursor to the beginning
  if (rsWaiver.CursorType > 0) {
    if (!rsWaiver.BOF) rsWaiver.MoveFirst();
  } else {
    rsWaiver.Requery();
  }

  // set the number of rows displayed on this page
  if (rsWaiver_numRows < 0 || rsWaiver_numRows > rsWaiver_total) {
    rsWaiver_numRows = rsWaiver_total;
  }

  // set the first and last displayed record
  rsWaiver_last  = Math.min(rsWaiver_first + rsWaiver_numRows - 1, rsWaiver_total);
  rsWaiver_first = Math.min(rsWaiver_first, rsWaiver_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsWaiver;
var MM_rsCount   = rsWaiver_total;
var MM_size      = rsWaiver_numRows;
var MM_uniqueCol = "";
    MM_paramName = "";
var MM_offset = 0;
var MM_atTotal = false;
var MM_paramIsDefined = (MM_paramName != "" && String(Request(MM_paramName)) != "undefined");
%>
<%
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
%>
<%
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
%>
<%
// *** Move To Record: update recordset stats

// set the first and last displayed record
rsWaiver_first = MM_offset + 1;
rsWaiver_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsWaiver_first = Math.min(rsWaiver_first, MM_rsCount);
  rsWaiver_last  = Math.min(rsWaiver_last, MM_rsCount);
}

// set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount != -1 && MM_offset + MM_size >= MM_rsCount);
%>
<%
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
%>
<%
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
	<title>Waiver Browse</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=290,height=150,scrollbars=1,left=0,top=0,status=1");
		return ;
	}		
	</Script>
</head>
<body>
<h5>Waiver</h5>
<table cellspacing="1">
    <tr> 
		<td><% if (MM_offset != 0) { %> <a href="<%=MM_moveFirst%>">First Page</a> <% } else { %> First Page <% } // end MM_offset != 0 %> |</td>
		<td><% if (MM_offset != 0) { %>	<a href="<%=MM_movePrev%>">Previous Page</a> <% } else { %> Previous Page <% } // end MM_offset != 0 %>|</td>
		<td><% if (!MM_atTotal) { %> <a href="<%=MM_moveNext%>">Next Page</a> <% } else { %> Next Page <% } // end !MM_atTotal %> |</td>
		<td><% if (!MM_atTotal) { %> <a href="<%=MM_moveLast%>">Last Page</a> <% } else { %> Last Page	<% } // end !MM_atTotal %></td>
	</tr>
	<tr>
		<td colspan="4">Displaying Records <%=rsWaiver_first%> To <%=rsWaiver_last%> Of <%=rsWaiver_total%></td>
	</tr>	
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
    <tr> 
		<th class="headrow" align="left">Date Received</th>	
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsWaiver.EOF)) { 
%>	
    <tr> 
		<td><a href="m001e0604.asp?intWaiver_id=<%=(rsWaiver.Fields.Item("intWaiver_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>"><%=FilterDate(rsWaiver.Fields.Item("dtsWaiverDate").Value)%></a>&nbsp;</td>	
    </tr>
<%  
	Repeat1__index++;
	rsWaiver.MoveNext();
} 
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><a href="javascript: openWindow('m001a0604.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>','wA0604');">Add Waiver</a></td>
	</tr>
</table>
</body>
</html>
<%
rsWaiver.Close();
%>
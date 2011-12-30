<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsLoanHistory = Server.CreateObject("ADODB.Recordset");
rsLoanHistory.ActiveConnection = MM_cnnASP02_STRING;
rsLoanHistory.Source = "{call dbo.cp_Get_Ivtry_Loan("+ Request.QueryString("intEquip_Set_id") + ",0,0)}";
rsLoanHistory.CursorType = 0;
rsLoanHistory.CursorLocation = 2;
rsLoanHistory.LockType = 3;
rsLoanHistory.Open();
var rsLoanHistory_numRows = 0;
var Repeat1__numRows = 5;
var Repeat1__index = 0;
rsLoanHistory_numRows += Repeat1__numRows;
// set the record count
var rsLoanHistory_total = rsLoanHistory.RecordCount;

// set the number of rows displayed on this page
if (rsLoanHistory_numRows < 0) {            // if repeat region set to all records
  rsLoanHistory_numRows = rsLoanHistory_total;
} else if (rsLoanHistory_numRows == 0) {    // if no repeat regions
  rsLoanHistory_numRows = 1;
}

// set the first and last displayed record
var rsLoanHistory_first = 1;
var rsLoanHistory_last  = rsLoanHistory_first + rsLoanHistory_numRows - 1;

// if we have the correct record count, check the other stats
if (rsLoanHistory_total != -1) {
  rsLoanHistory_numRows = Math.min(rsLoanHistory_numRows, rsLoanHistory_total);
  rsLoanHistory_first   = Math.min(rsLoanHistory_first, rsLoanHistory_total);
  rsLoanHistory_last    = Math.min(rsLoanHistory_last, rsLoanHistory_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsLoanHistory_total == -1) {

  // count the total records by iterating through the recordset
  for (rsLoanHistory_total=0; !rsLoanHistory.EOF; rsLoanHistory.MoveNext()) {
    rsLoanHistory_total++;
  }

  // reset the cursor to the beginning
  if (rsLoanHistory.CursorType > 0) {
    if (!rsLoanHistory.BOF) rsLoanHistory.MoveFirst();
  } else {
    rsLoanHistory.Requery();
  }

  // set the number of rows displayed on this page
  if (rsLoanHistory_numRows < 0 || rsLoanHistory_numRows > rsLoanHistory_total) {
    rsLoanHistory_numRows = rsLoanHistory_total;
  }

  // set the first and last displayed record
  rsLoanHistory_last  = Math.min(rsLoanHistory_first + rsLoanHistory_numRows - 1, rsLoanHistory_total);
  rsLoanHistory_first = Math.min(rsLoanHistory_first, rsLoanHistory_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsLoanHistory;
var MM_rsCount   = rsLoanHistory_total;
var MM_size      = rsLoanHistory_numRows;
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
rsLoanHistory_first = MM_offset + 1;
rsLoanHistory_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsLoanHistory_first = Math.min(rsLoanHistory_first, MM_rsCount);
  rsLoanHistory_last  = Math.min(rsLoanHistory_last, MM_rsCount);
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

// set the record count
var rsLoanHistory_total = rsLoanHistory.RecordCount;

// set the number of rows displayed on this page
if (rsLoanHistory_numRows < 0) {            // if repeat region set to all records
  rsLoanHistory_numRows = rsLoanHistory_total;
} else if (rsLoanHistory_numRows == 0) {    // if no repeat regions
  rsLoanHistory_numRows = 1;
}

// set the first and last displayed record
var rsLoanHistory_first = 1;
var rsLoanHistory_last  = rsLoanHistory_first + rsLoanHistory_numRows - 1;

// if we have the correct record count, check the other stats
if (rsLoanHistory_total != -1) {
  rsLoanHistory_numRows = Math.min(rsLoanHistory_numRows, rsLoanHistory_total);
  rsLoanHistory_first   = Math.min(rsLoanHistory_first, rsLoanHistory_total);
  rsLoanHistory_last    = Math.min(rsLoanHistory_last, rsLoanHistory_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsLoanHistory_total == -1) {

  // count the total records by iterating through the recordset
  for (rsLoanHistory_total=0; !rsLoanHistory.EOF; rsLoanHistory.MoveNext()) {
    rsLoanHistory_total++;
  }

  // reset the cursor to the beginning
  if (rsLoanHistory.CursorType > 0) {
    if (!rsLoanHistory.BOF) rsLoanHistory.MoveFirst();
  } else {
    rsLoanHistory.Requery();
  }

  // set the number of rows displayed on this page
  if (rsLoanHistory_numRows < 0 || rsLoanHistory_numRows > rsLoanHistory_total) {
    rsLoanHistory_numRows = rsLoanHistory_total;
  }

  // set the first and last displayed record
  rsLoanHistory_last  = Math.min(rsLoanHistory_first + rsLoanHistory_numRows - 1, rsLoanHistory_total);
  rsLoanHistory_first = Math.min(rsLoanHistory_first, rsLoanHistory_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsLoanHistory;
var MM_rsCount   = rsLoanHistory_total;
var MM_size      = rsLoanHistory_numRows;
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
rsLoanHistory_first = MM_offset + 1;
rsLoanHistory_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsLoanHistory_first = Math.min(rsLoanHistory_first, MM_rsCount);
  rsLoanHistory_last  = Math.min(rsLoanHistory_last, MM_rsCount);
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

// set the record count
var rsLoanHistory_total = rsLoanHistory.RecordCount;

// set the number of rows displayed on this page
if (rsLoanHistory_numRows < 0) {            // if repeat region set to all records
  rsLoanHistory_numRows = rsLoanHistory_total;
} else if (rsLoanHistory_numRows == 0) {    // if no repeat regions
  rsLoanHistory_numRows = 1;
}

// set the first and last displayed record
var rsLoanHistory_first = 1;
var rsLoanHistory_last  = rsLoanHistory_first + rsLoanHistory_numRows - 1;

// if we have the correct record count, check the other stats
if (rsLoanHistory_total != -1) {
  rsLoanHistory_numRows = Math.min(rsLoanHistory_numRows, rsLoanHistory_total);
  rsLoanHistory_first   = Math.min(rsLoanHistory_first, rsLoanHistory_total);
  rsLoanHistory_last    = Math.min(rsLoanHistory_last, rsLoanHistory_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsLoanHistory_total == -1) {

  // count the total records by iterating through the recordset
  for (rsLoanHistory_total=0; !rsLoanHistory.EOF; rsLoanHistory.MoveNext()) {
    rsLoanHistory_total++;
  }

  // reset the cursor to the beginning
  if (rsLoanHistory.CursorType > 0) {
    if (!rsLoanHistory.BOF) rsLoanHistory.MoveFirst();
  } else {
    rsLoanHistory.Requery();
  }

  // set the number of rows displayed on this page
  if (rsLoanHistory_numRows < 0 || rsLoanHistory_numRows > rsLoanHistory_total) {
    rsLoanHistory_numRows = rsLoanHistory_total;
  }

  // set the first and last displayed record
  rsLoanHistory_last  = Math.min(rsLoanHistory_first + rsLoanHistory_numRows - 1, rsLoanHistory_total);
  rsLoanHistory_first = Math.min(rsLoanHistory_first, rsLoanHistory_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsLoanHistory;
var MM_rsCount   = rsLoanHistory_total;
var MM_size      = rsLoanHistory_numRows;
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
  for (var i=0; !rsLoanHistory.EOF && (i < MM_offset || MM_offset == -1); i++) {
    rsLoanHistory.MoveNext();
  }
  if (rsLoanHistory.EOF) MM_offset = i;  // set MM_offset to the last possible record
}

// *** Move To Record: if we dont know the record count, check the display range

if (MM_rsCount == -1) {

  // walk to the end of the display range for this page
  for (var i=MM_offset; !rsLoanHistory.EOF && (MM_size < 0 || i < MM_offset + MM_size); i++) {
    rsLoanHistory.MoveNext();
  }

  // if we walked off the end of the recordset, set MM_rsCount and MM_size
  if (rsLoanHistory.EOF) {
    MM_rsCount = i;
    if (MM_size < 0 || MM_size > MM_rsCount) MM_size = MM_rsCount;
  }

  // if we walked off the end, set the offset based on page size
  if (rsLoanHistory.EOF && !MM_paramIsDefined) {
    if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
      MM_offset = MM_rsCount - (MM_rsCount % MM_size);
    } else {
      MM_offset = MM_rsCount - MM_size;
    }
  }

  // reset the cursor to the beginning
  if (rsLoanHistory.CursorType > 0) {
    if (!MM_rs.BOF) rsLoanHistory.MoveFirst();
  } else {
    rsLoanHistory.Requery();
  }

  // move the cursor to the selected record
  for (var i=0; !rsLoanHistory.EOF && i < MM_offset; i++) {
    rsLoanHistory.MoveNext();
  }
}
// *** Move To Record: update recordset stats

// set the first and last displayed record
rsLoanHistory_first = MM_offset + 1;
rsLoanHistory_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsLoanHistory_first = Math.min(rsLoanHistory_first, MM_rsCount);
  rsLoanHistory_last  = Math.min(rsLoanHistory_last, MM_rsCount);
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
	<title>Loan History</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</script>	
</head>
<body>
<h5>Loan History</h5>
<table cellspacing="1">
	<tr>
    	<td><% if (MM_offset != 0) { %><a href="<%=MM_moveFirst%>">First Page</a><% } else { %>First Page<%	} // end MM_offset != 0 %>|</td>
		<td><% if (MM_offset != 0) { %><a href="<%=MM_movePrev%>">Previous Page</a><% } else { %>Previous Page <% } // end MM_offset != 0 %>|</td>
		<td><% if (!MM_atTotal) { %><a href="<%=MM_moveNext%>">Next Page</a><% } else { %>Next Page <% } // end !MM_atTotal %>|</td>
		<td><% if (!MM_atTotal) { %><a href="<%=MM_moveLast%>">Last Page</a><% } else { %>Last Page <% } // end !MM_atTotal %></td>
    </tr>
    <tr>
		<td colspan="4" align="left">Displaying Records <b><%=(rsLoanHistory_first)%></b> to <b><%=(rsLoanHistory_last)%></b> of <b><%=(rsLoanHistory_total)%></b></td>
    </tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="MTable">
	<tr>
		<th class="headrow" nowrap align="left">Loan Request</th>
		<th class="headrow" nowrap align="left">Loaned To</th>		
		<th class="headrow" nowrap align="left">Loan Status</th>
		<th class="headrow" nowrap align="left">Loan Type</th>
		<th class="headrow" nowrap align="center">Date Processed</th>
		<th class="headrow" nowrap align="center">Delivery Date</th>
		<th class="headrow" nowrap align="center">Return Date</th>
		<th class="headrow" nowrap align="left">Returned By</th>
	</tr>
<%
while ((Repeat1__numRows-- != 0) && (!rsLoanHistory.EOF)) {
%>
	<tr>
		<td nowrap align="left"><a href="javascript: openWindow('../LN/m008FS3.asp?intLoan_req_id=<%=rsLoanHistory.Fields.Item("intLoan_Req_id").Value%>','');"><%=ZeroPadFormat(rsLoanHistory.Fields.Item("intLoan_Req_id").Value,8)%></a>&nbsp;</td>
		<td nowrap align="left"><%=(rsLoanHistory.Fields.Item("chvLoaned_to").Value)%>&nbsp;</td>		
		<td nowrap align="left"><%=(rsLoanHistory.Fields.Item("chvLoan_Status").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsLoanHistory.Fields.Item("chvLoan_Type").Value)%>&nbsp;</td>
		<td nowrap align="center"><%=FilterDate(rsLoanHistory.Fields.Item("dtsDate_Shipped").Value)%></td>
		<td nowrap align="center"><%=FilterDate(rsLoanHistory.Fields.Item("dtsDlvy_date").Value)%>&nbsp;</td>
		<td nowrap align="center"><%=FilterDate(rsLoanHistory.Fields.Item("dtsDate_Returned").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsLoanHistory.Fields.Item("chvReturned_by").Value)%>&nbsp;</td>
	</tr>
<%
	Repeat1__index++;
	rsLoanHistory.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsLoanHistory.Close();
%>
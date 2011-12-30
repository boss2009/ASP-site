<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.QueryString("Remove"))!="undefined") {
	var rsRemoveMailList = Server.CreateObject("ADODB.Recordset");
	rsRemoveMailList.ActiveConnection = MM_cnnASP02_STRING;
	rsRemoveMailList.Source = "{call dbo.cp_ctc_mail_list("+Request.QueryString("intCtm_Id")+",0,0,0,'D',0)}";
	rsRemoveMailList.CursorType = 0;
	rsRemoveMailList.CursorLocation = 2;
	rsRemoveMailList.LockType = 3;
	rsRemoveMailList.Open();
}

var rsMailList = Server.CreateObject("ADODB.Recordset");
rsMailList.ActiveConnection = MM_cnnASP02_STRING;
rsMailList.Source = "{call dbo.cp_ctc_mail_list(0,"+Request.QueryString("intContact_id")+",0,0,'Q',0)}";
rsMailList.CursorType = 0;
rsMailList.CursorLocation = 2;
rsMailList.LockType = 3;
rsMailList.Open();
var rsMailList_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsMailList_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsMailList_total = rsMailList.RecordCount;

// set the number of rows displayed on this page
if (rsMailList_numRows < 0) {            // if repeat region set to all records
  rsMailList_numRows = rsMailList_total;
} else if (rsMailList_numRows == 0) {    // if no repeat regions
  rsMailList_numRows = 1;
}

// set the first and last displayed record
var rsMailList_first = 1;
var rsMailList_last  = rsMailList_first + rsMailList_numRows - 1;

// if we have the correct record count, check the other stats
if (rsMailList_total != -1) {
  rsMailList_numRows = Math.min(rsMailList_numRows, rsMailList_total);
  rsMailList_first   = Math.min(rsMailList_first, rsMailList_total);
  rsMailList_last    = Math.min(rsMailList_last, rsMailList_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsMailList_total == -1) {

  // count the total records by iterating through the recordset
  for (rsMailList_total=0; !rsMailList.EOF; rsMailList.MoveNext()) {
    rsMailList_total++;
  }

  // reset the cursor to the beginning
  if (rsMailList.CursorType > 0) {
    if (!rsMailList.BOF) rsMailList.MoveFirst();
  } else {
    rsMailList.Requery();
  }

  // set the number of rows displayed on this page
  if (rsMailList_numRows < 0 || rsMailList_numRows > rsMailList_total) {
    rsMailList_numRows = rsMailList_total;
  }

  // set the first and last displayed record
  rsMailList_last  = Math.min(rsMailList_first + rsMailList_numRows - 1, rsMailList_total);
  rsMailList_first = Math.min(rsMailList_first, rsMailList_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsMailList;
var MM_rsCount   = rsMailList_total;
var MM_size      = rsMailList_numRows;
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
rsMailList_first = MM_offset + 1;
rsMailList_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsMailList_first = Math.min(rsMailList_first, MM_rsCount);
  rsMailList_last  = Math.min(rsMailList_last, MM_rsCount);
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
	<title>Mailing List</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
		if (window.focus) self.focus();
		function openWindow(page, name){
			if (page!='nothing') win1=window.open(page, "", "width=400,height=200,scrollbars=1,left=0,top=0,status=1");
			return ;
		}
	</Script>
</head>
<body>
<h5>Mailing List</h5>
<table cellspacing="1">
    <tr>
		<td><% if (MM_offset != 0) { %> <a href="<%=MM_moveFirst%>">First Page</a> <% } else { %> First Page <%	} // end MM_offset != 0 %> |</td>
		<td><% if (MM_offset != 0) { %>	<a href="<%=MM_movePrev%>">Previous Page</a> <% } else { %>	Previous Page <% } // end MM_offset != 0 %> |</td>
		<td><% if (!MM_atTotal) { %> <a href="<%=MM_moveNext%>">Next Page</a> <% } else { %> Next Page <% } // end !MM_atTotal %> |</td>
		<td><% if (!MM_atTotal) { %> <a href="<%=MM_moveLast%>">Last Page</a> <% } else { %> Last Page <% } // end !MM_atTotal %></td>
    </tr>
    <tr>
		<td colspan="4" align="left">Displaying Records <b><%=(rsMailList_first)%></b> to <b><%=(rsMailList_last)%></b> of <b><%=(rsMailList_total)%></b></td>
    </tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr> 
		<th nowrap class="headrow" align="left" width="300">Name</th>
		<th nowrap class="headrow" align="left">&nbsp;</th>
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsMailList.EOF)) { 
%>
    <tr> 
		<td nowrap><%=rsMailList.Fields.Item("chvName").Value%></td>
		<td nowrap><a href="m004q0301.asp?intContact_id=<%=Request.QueryString("intContact_id")%>&Remove=1&intCtm_Id=<%=rsMailList.Fields.Item("intCtm_Id").Value%>">Unsubscribe</a></td>
    </tr>
<%
	Repeat1__index++;
	rsMailList.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><a href="javascript: openWindow('m004a0301.asp?intContact_id=<%=Request.QueryString("intContact_id")%>','w004A03');">Subscribe Mailing List</a></td>
	</tr>
</table>
</body>
</html>
<%
rsMailList.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsOnsiteSupport = Server.CreateObject("ADODB.Recordset");
rsOnsiteSupport.ActiveConnection = MM_cnnASP02_STRING;
rsOnsiteSupport.Source = "{call dbo.cp_pilat_site_support("+ Request.QueryString("intReferral_id") + ",0,0,'Q',0)}";
rsOnsiteSupport.CursorType = 0;
rsOnsiteSupport.CursorLocation = 2;
rsOnsiteSupport.LockType = 3;
rsOnsiteSupport.Open();
var rsOnsiteSupport_numRows = 0;
%>
<%
var Repeat1__numRows = 5;
var Repeat1__index = 0;
rsOnsiteSupport_numRows += Repeat1__numRows;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsOnsiteSupport_total = rsOnsiteSupport.RecordCount;

// set the number of rows displayed on this page
if (rsOnsiteSupport_numRows < 0) {            // if repeat region set to all records
  rsOnsiteSupport_numRows = rsOnsiteSupport_total;
} else if (rsOnsiteSupport_numRows == 0) {    // if no repeat regions
  rsOnsiteSupport_numRows = 1;
}

// set the first and last displayed record
var rsOnsiteSupport_first = 1;
var rsOnsiteSupport_last  = rsOnsiteSupport_first + rsOnsiteSupport_numRows - 1;

// if we have the correct record count, check the other stats
if (rsOnsiteSupport_total != -1) {
  rsOnsiteSupport_numRows = Math.min(rsOnsiteSupport_numRows, rsOnsiteSupport_total);
  rsOnsiteSupport_first   = Math.min(rsOnsiteSupport_first, rsOnsiteSupport_total);
  rsOnsiteSupport_last    = Math.min(rsOnsiteSupport_last, rsOnsiteSupport_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsOnsiteSupport_total == -1) {

  // count the total records by iterating through the recordset
  for (rsOnsiteSupport_total=0; !rsOnsiteSupport.EOF; rsOnsiteSupport.MoveNext()) {
    rsOnsiteSupport_total++;
  }

  // reset the cursor to the beginning
  if (rsOnsiteSupport.CursorType > 0) {
    if (!rsOnsiteSupport.BOF) rsOnsiteSupport.MoveFirst();
  } else {
    rsOnsiteSupport.Requery();
  }

  // set the number of rows displayed on this page
  if (rsOnsiteSupport_numRows < 0 || rsOnsiteSupport_numRows > rsOnsiteSupport_total) {
    rsOnsiteSupport_numRows = rsOnsiteSupport_total;
  }

  // set the first and last displayed record
  rsOnsiteSupport_last  = Math.min(rsOnsiteSupport_first + rsOnsiteSupport_numRows - 1, rsOnsiteSupport_total);
  rsOnsiteSupport_first = Math.min(rsOnsiteSupport_first, rsOnsiteSupport_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsOnsiteSupport;
var MM_rsCount   = rsOnsiteSupport_total;
var MM_size      = rsOnsiteSupport_numRows;
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
rsOnsiteSupport_first = MM_offset + 1;
rsOnsiteSupport_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsOnsiteSupport_first = Math.min(rsOnsiteSupport_first, MM_rsCount);
  rsOnsiteSupport_last  = Math.min(rsOnsiteSupport_last, MM_rsCount);
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
	<title>On-site Support</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=650,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</Script>	
</head>
<body>
<h5>On-site Support</h5>
<table cellspacing="1">
	<tr>
		<td><% if (MM_offset != 0) { %><a href="<%=MM_moveFirst%>">First Page</a><% } else { %>First Page<% } // end MM_offset != 0 %>|</td>
		<td><% if (MM_offset != 0) { %><a href="<%=MM_movePrev%>">Previous Page</a><% } else { %>Previous Page<% } // end MM_offset != 0 %>|</td>
		<td><% if (!MM_atTotal) { %><a href="<%=MM_moveNext%>">Next Page</a><% } else { %>Next Page<% } // end !MM_atTotal %>|</td>
		<td><% if (!MM_atTotal) { %><a href="<%=MM_moveLast%>">Last Page</a><% } else { %>Last Page<% } // end !MM_atTotal %></td>
	</tr>
	<tr>
		<td colspan="4">Displaying Records <%=rsOnsiteSupport_first%> To <%=rsOnsiteSupport_last%> Of <%=rsOnsiteSupport_total%></td>
	</tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr> 
		<th class="headrow" nowrap align="left" width="180">Name</th>
		<th class="headrow">&nbsp;</th>			
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsOnsiteSupport.EOF)) { 
%>
    <tr> 
		<td nowrap><a href="javascript: openWindow('../CT/m004FS3.asp?intContact_id=<%=(rsOnsiteSupport.Fields.Item("intContact_id").Value)%>&intReferral_id=<%=Request.QueryString("intReferral_id")%>');"><%=rsOnsiteSupport.Fields.Item("chvContact").Value%></a>&nbsp;</td>
		<td nowrap><a href="javascript: openWindow('m012q0204x.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&intContact_id=<%=(rsOnsiteSupport.Fields.Item("intContact_id").Value)%>');"><img src="../i/remove.gif" ALT="Remove <%=rsOnsiteSupport.Fields.Item("chvContact").Value%>"></a></td>		
    </tr>
<%
	Repeat1__index++;
	rsOnsiteSupport.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><a href="javascript: openWindow('../CT/m004a0101.asp?LinkToClass=4&LinkToObject=<%=Request.QueryString("intReferral_id")%>','winAdd');">Add On-site Support</a></td>
	</tr>
</table>
</body>
</html>
<%
rsOnsiteSupport.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var rsInventoryStatus = Server.CreateObject("ADODB.Recordset");
rsInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryStatus.Source = "{call dbo.cp_Get_Equip_Status(0,0)}";
rsInventoryStatus.CursorType = 0;
rsInventoryStatus.CursorLocation = 2;
rsInventoryStatus.LockType = 3;
rsInventoryStatus.Open();
var rsInventoryStatus_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsInventoryStatus_numRows += Repeat1__numRows;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsInventoryStatus_total = rsInventoryStatus.RecordCount;

// set the number of rows displayed on this page
if (rsInventoryStatus_numRows < 0) {            // if repeat region set to all records
  rsInventoryStatus_numRows = rsInventoryStatus_total;
} else if (rsInventoryStatus_numRows == 0) {    // if no repeat regions
  rsInventoryStatus_numRows = 1;
}

// set the first and last displayed record
var rsInventoryStatus_first = 1;
var rsInventoryStatus_last  = rsInventoryStatus_first + rsInventoryStatus_numRows - 1;

// if we have the correct record count, check the other stats
if (rsInventoryStatus_total != -1) {
  rsInventoryStatus_numRows = Math.min(rsInventoryStatus_numRows, rsInventoryStatus_total);
  rsInventoryStatus_first   = Math.min(rsInventoryStatus_first, rsInventoryStatus_total);
  rsInventoryStatus_last    = Math.min(rsInventoryStatus_last, rsInventoryStatus_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsInventoryStatus_total == -1) {

  // count the total records by iterating through the recordset
  for (rsInventoryStatus_total=0; !rsInventoryStatus.EOF; rsInventoryStatus.MoveNext()) {
    rsInventoryStatus_total++;
  }

  // reset the cursor to the beginning
  if (rsInventoryStatus.CursorType > 0) {
    if (!rsInventoryStatus.BOF) rsInventoryStatus.MoveFirst();
  } else {
    rsInventoryStatus.Requery();
  }

  // set the number of rows displayed on this page
  if (rsInventoryStatus_numRows < 0 || rsInventoryStatus_numRows > rsInventoryStatus_total) {
    rsInventoryStatus_numRows = rsInventoryStatus_total;
  }

  // set the first and last displayed record
  rsInventoryStatus_last  = Math.min(rsInventoryStatus_first + rsInventoryStatus_numRows - 1, rsInventoryStatus_total);
  rsInventoryStatus_first = Math.min(rsInventoryStatus_first, rsInventoryStatus_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsInventoryStatus;
var MM_rsCount   = rsInventoryStatus_total;
var MM_size      = rsInventoryStatus_numRows;
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
rsInventoryStatus_first = MM_offset + 1;
rsInventoryStatus_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsInventoryStatus_first = Math.min(rsInventoryStatus_first, MM_rsCount);
  rsInventoryStatus_last  = Math.min(rsInventoryStatus_last, MM_rsCount);
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
	<title>Inventory Status Lookup Table</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=450,height=450,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</Script>
</head>
<body>
<h5>Inventory Status Lookup Table</h5>
<a href="../../aspMenu.asp">Master Menu</a> / <a href="../m018Menu.asp">Administrative Options</a> / <a href="../m018Sm03.asp">System Lookup Tables</a>
<table cellspacing="1">
	<tr>
		<td><% if (MM_offset != 0) { %><a href="<%=MM_moveFirst%>">First Page</a><% } else { %>First Page<% } // end MM_offset != 0 %>|</td>
		<td><% if (MM_offset != 0) { %><a href="<%=MM_movePrev%>">Previous Page</a><% } else { %>Previous Page<% } // end MM_offset != 0 %>|</td>
		<td><% if (!MM_atTotal) { %><a href="<%=MM_moveNext%>">Next Page</a><% } else { %>Next Page<% } // end !MM_atTotal %>|</td>
		<td><% if (!MM_atTotal) { %><a href="<%=MM_moveLast%>">Last Page</a><% } else { %>Last Page<% } // end !MM_atTotal %></td>
	</tr>
</table>    
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
    <tr> 
		<th class="headrow" align="left">Description</th>
		<th class="headrow" align="left">Is Active</th>
		<th class="headrow" align="left">Is Manual Select</th>
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsInventoryStatus.EOF)) { 
%>
    <tr> 
		<td><a href="m018e0336.asp?insEquip_status_id=<%=(rsInventoryStatus.Fields.Item("insEquip_status_id").Value)%>"><%=(rsInventoryStatus.Fields.Item("chvStatusDesc").Value)%></a></td>
		<td align="center"><%=(rsInventoryStatus.Fields.Item("bitis_active").Value)%></td>
		<td align="center"><%=(rsInventoryStatus.Fields.Item("bitis_manselect").Value)%></td>		
    </tr>
<%
	Repeat1__index++;
	rsInventoryStatus.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><a href="javascript: openWindow('m018a0336.asp','w18A0336');">Add Inventory Status</a></td>
	</tr>
</table>
</body>
</html>
<%
rsInventoryStatus.Close();
%>

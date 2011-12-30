<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" --> 
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var rsLetterTemplate = Server.CreateObject("ADODB.Recordset");
rsLetterTemplate.ActiveConnection = MM_cnnASP02_STRING;
rsLetterTemplate.Source = "{call dbo.cp_Letter_template(0,0,'',0,'',0,0,0,0,0,0,0,'Q',0)}";
rsLetterTemplate.CursorType = 0;
rsLetterTemplate.CursorLocation = 2;
rsLetterTemplate.LockType = 3;
rsLetterTemplate.Open();
var rsLetterTemplate_numRows = 0;
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsLetterTemplate_numRows += Repeat1__numRows;
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsLetterTemplate_total = rsLetterTemplate.RecordCount;

// set the number of rows displayed on this page
if (rsLetterTemplate_numRows < 0) {            // if repeat region set to all records
  rsLetterTemplate_numRows = rsLetterTemplate_total;
} else if (rsLetterTemplate_numRows == 0) {    // if no repeat regions
  rsLetterTemplate_numRows = 1;
}

// set the first and last displayed record
var rsLetterTemplate_first = 1;
var rsLetterTemplate_last  = rsLetterTemplate_first + rsLetterTemplate_numRows - 1;

// if we have the correct record count, check the other stats
if (rsLetterTemplate_total != -1) {
  rsLetterTemplate_numRows = Math.min(rsLetterTemplate_numRows, rsLetterTemplate_total);
  rsLetterTemplate_first   = Math.min(rsLetterTemplate_first, rsLetterTemplate_total);
  rsLetterTemplate_last    = Math.min(rsLetterTemplate_last, rsLetterTemplate_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsLetterTemplate_total == -1) {

  // count the total records by iterating through the recordset
  for (rsLetterTemplate_total=0; !rsLetterTemplate.EOF; rsLetterTemplate.MoveNext()) {
    rsLetterTemplate_total++;
  }

  // reset the cursor to the beginning
  if (rsLetterTemplate.CursorType > 0) {
    if (!rsLetterTemplate.BOF) rsLetterTemplate.MoveFirst();
  } else {
    rsLetterTemplate.Requery();
  }

  // set the number of rows displayed on this page
  if (rsLetterTemplate_numRows < 0 || rsLetterTemplate_numRows > rsLetterTemplate_total) {
    rsLetterTemplate_numRows = rsLetterTemplate_total;
  }

  // set the first and last displayed record
  rsLetterTemplate_last  = Math.min(rsLetterTemplate_first + rsLetterTemplate_numRows - 1, rsLetterTemplate_total);
  rsLetterTemplate_first = Math.min(rsLetterTemplate_first, rsLetterTemplate_total);
}

var MM_paramName = "";
var MM_rs        = rsLetterTemplate;
var MM_rsCount   = rsLetterTemplate_total;
var MM_size      = rsLetterTemplate_numRows;
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
rsLetterTemplate_first = MM_offset + 1;
rsLetterTemplate_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsLetterTemplate_first = Math.min(rsLetterTemplate_first, MM_rsCount);
  rsLetterTemplate_last  = Math.min(rsLetterTemplate_last, MM_rsCount);
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
	<title>Letter Template Lookup Table</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=350,height=350,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>
</head>
<body>
<h5>Letter Template Lookup Table</h5>
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
		<th class="headrow" align="left">Template Name</th>
		<th class="headrow" align="left">Template Type</th>
		<th class="headrow" align="left">Document Type</th>		
		<th class="headrow" align="left">File Name</th>
		<th class="headrow" align="center">Loan Document</th>		
		<th class="headrow" align="center">Outstanding Document</th>		
		<th class="headrow" align="center">Decline Document</th>		
		<th class="headrow" align="center">Pending Document</th>		
		<th class="headrow" align="center">Include Equipment</th>		
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsLetterTemplate.EOF)) { 
%>
    <tr> 		
		<td><a href="m018e0341.asp?insTemplate_id=<%=(rsLetterTemplate.Fields.Item("insTemplate_id").Value)%>"><%=(rsLetterTemplate.Fields.Item("chvTemplate_Name").Value)%></a>&nbsp;</td>				
		<td align="center"><%=(rsLetterTemplate.Fields.Item("chvTemplate_Type").Value)%>&nbsp;</td>		
		<td align="center"><%=(rsLetterTemplate.Fields.Item("chvDocType").Value)%>&nbsp;</td>		
		<td align="left"><%=(rsLetterTemplate.Fields.Item("chvFileName").Value)%>&nbsp;</td>
		<td align="center"><%=(rsLetterTemplate.Fields.Item("bitIs_Loan_Doc").Value)%>&nbsp;</td>		
		<td align="center"><%=(rsLetterTemplate.Fields.Item("bitIs_OutStand_Doc").Value)%>&nbsp;</td>		
		<td align="center"><%=(rsLetterTemplate.Fields.Item("bitIs_Decline_Doc").Value)%>&nbsp;</td>		
		<td align="center"><%=(rsLetterTemplate.Fields.Item("bitIs_Pending_Doc").Value)%>&nbsp;</td>		
		<td align="center"><%=(rsLetterTemplate.Fields.Item("bitIs_Include_Eqp").Value)%>&nbsp;</td>		
    </tr>
<%  
	Repeat1__index++;
	rsLetterTemplate.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><a href="javascript: openWindow('m018a0341.asp','w18A0341');">Add Letter Template</a></td>
	</tr>
</table>
</body>
</html>
<%
rsLetterTemplate.Close();
%>
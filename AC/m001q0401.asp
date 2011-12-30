<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
if (String(Request.QueryString("Remove"))=="True") {
	var rsRemoveContact = Server.CreateObject("ADODB.Recordset");
	rsRemoveContact.ActiveConnection = MM_cnnASP02_STRING;
	rsRemoveContact.Source = "{call dbo.cp_ClnCtact2("+ Request.QueryString("intAdult_id") + ","+Request.QueryString("intContact_id")+",0,0,0,'D',0)}";
	rsRemoveContact.CursorType = 0;
	rsRemoveContact.CursorLocation = 2;
	rsRemoveContact.LockType = 3;
	rsRemoveContact.Open();	
	Response.Redirect("m001q0401.asp?intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_ClnCtact2("+ Request.QueryString("intAdult_id") + ",0,0,0,0,'Q',0)}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();
var rsContact_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsContact_numRows += Repeat1__numRows;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsContact_total = rsContact.RecordCount;

// set the number of rows displayed on this page
if (rsContact_numRows < 0) {            // if repeat region set to all records
  rsContact_numRows = rsContact_total;
} else if (rsContact_numRows == 0) {    // if no repeat regions
  rsContact_numRows = 1;
}

// set the first and last displayed record
var rsContact_first = 1;
var rsContact_last  = rsContact_first + rsContact_numRows - 1;

// if we have the correct record count, check the other stats
if (rsContact_total != -1) {
  rsContact_numRows = Math.min(rsContact_numRows, rsContact_total);
  rsContact_first   = Math.min(rsContact_first, rsContact_total);
  rsContact_last    = Math.min(rsContact_last, rsContact_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsContact_total == -1) {

  // count the total records by iterating through the recordset
  for (rsContact_total=0; !rsContact.EOF; rsContact.MoveNext()) {
    rsContact_total++;
  }

  // reset the cursor to the beginning
  if (rsContact.CursorType > 0) {
    if (!rsContact.BOF) rsContact.MoveFirst();
  } else {
    rsContact.Requery();
  }

  // set the number of rows displayed on this page
  if (rsContact_numRows < 0 || rsContact_numRows > rsContact_total) {
    rsContact_numRows = rsContact_total;
  }

  // set the first and last displayed record
  rsContact_last  = Math.min(rsContact_first + rsContact_numRows - 1, rsContact_total);
  rsContact_first = Math.min(rsContact_first, rsContact_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsContact;
var MM_rsCount   = rsContact_total;
var MM_size      = rsContact_numRows;
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
rsContact_first = MM_offset + 1;
rsContact_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsContact_first = Math.min(rsContact_first, MM_rsCount);
  rsContact_last  = Math.min(rsContact_last, MM_rsCount);
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
	<title>Client Contacts</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</Script>	
</head>
<body>
<h5>Client Contacts</h5>
<table cellspacing="1">
	<tr>
		<td><% if (MM_offset != 0) { %><a href="<%=MM_moveFirst%>">First Page</a><% } else { %>First Page<% } // end MM_offset != 0 %>|</td>
		<td><% if (MM_offset != 0) { %><a href="<%=MM_movePrev%>">Previous Page</a><% } else { %>Previous Page<% } // end MM_offset != 0 %>|</td>
		<td><% if (!MM_atTotal) { %><a href="<%=MM_moveNext%>">Next Page</a><% } else { %>Next Page<% } // end !MM_atTotal %>|</td>
		<td><% if (!MM_atTotal) { %><a href="<%=MM_moveLast%>">Last Page</a><% } else { %>Last Page<% } // end !MM_atTotal %></td>
	</tr>
	<tr>
		<td colspan="4">Displaying Records <%=rsContact_first%> To <%=rsContact_last%> Of <%=rsContact_total%></td>
	</tr>
</table>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 
<!--	
		<th class="headrow" nowrap align="left">Entry Date</th>
-->
		<th class="headrow" nowrap align="left" width="180px">Name</th>	
		<th class="headrow" nowrap align="left">Relationship</th>
		<th class="headrow" nowrap align="left">Phone Number</th>
		<th class="headrow" nowrap align="left">Organization</th>		
		<th class="headrow" nowrap align="left">Job Title</th>		
		<th class="headrow" nowrap align="left">Organization Type</th>
		<th class="headrow">&nbsp;</th>		
    </tr>
<% 
function Lookup(id){
	var temp = "";
	switch (String(id)) {
		case "0":
			temp = "None";
		break;
		case "1":
			temp = "Home";
		break;
		case "2":
			temp = "Off";
		break
		case "3":
			temp = "Cell";
		break;
		case "4":
			temp = "Pger";
		break;
		case "5":
			temp = "Fax";
		break;
		case "6":
			temp = "Wk";
		break;
	}
	return temp;
}
while ((Repeat1__numRows-- != 0) && (!rsContact.EOF)) { 
%>
    <tr> 
<!--	
		<td nowrap><%=FilterDate(rsContact.Fields.Item("dtsEntrydate").Value)%>&nbsp;</td>			
-->
		<td nowrap><a href="javascript: openWindow('../CT/m004FS3.asp?intContact_id=<%=(rsContact.Fields.Item("intContact_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>');"><%=(rsContact.Fields.Item("chvContact_Name").Value)%></a>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("chvRelationship").Value)%>&nbsp;<a href="javascript: openWindow('../CT/m004e0102.asp?LinkToClass=1&LinkToObject=<%=Request.QueryString("intAdult_id")%>&intContact_id=<%=(rsContact.Fields.Item("intContact_id").Value)%>&ContactName=<%=FilterQuotes(rsContact.Fields.Item("chvContact_Name").Value)%>&ObjectName=<%=FilterQuotes(rsClient.Fields.Item("chvFst_Name").Value)%> <%=FilterQuotes(rsClient.Fields.Item("chvLst_Name").Value)%>&Relationship=<%=FilterQuotes(rsContact.Fields.Item("chvRelationship").Value)%>');">[change]</a></td>
<!--
// removed at Oct.28.2005
-->
		<td nowrap><%=(rsContact.Fields.Item("Inst_Company").Value)%>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("chvJob_title").Value)%>&nbsp;</td>				
		<td nowrap><%=(rsContact.Fields.Item("chvWork_Type").Value)%>&nbsp;</td>		
		<td nowrap><a href="m001q0401.asp?intContact_id=<%=(rsContact.Fields.Item("intContact_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>&Remove=True"><img src="../i/Remove.gif" ALT="Remove <%=(rsContact.Fields.Item("chvContact_Name").Value)%>"></a></td>		
    </tr>
<%
	Repeat1__index++;
	rsContact.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><a href="javascript: openWindow('../CT/m004a0101.asp?LinkToClass=1&LinkToObject=<%=Request.QueryString("intAdult_id")%>','winAdd');">Add Contact</a></td>
	</tr>
</table>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
%>
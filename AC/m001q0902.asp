<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsOldCorrespondence = Server.CreateObject("ADODB.Recordset");
rsOldCorrespondence.ActiveConnection = MM_cnnASP02_STRING;
rsOldCorrespondence.Source = "{call dbo.cp_Clnt_Crsp("+ Request.QueryString("intAdult_id") + ")}";
rsOldCorrespondence.CursorType = 0;
rsOldCorrespondence.CursorLocation = 2;
rsOldCorrespondence.LockType = 3;
rsOldCorrespondence.Open();
var rsOldCorrespondence_numRows = 0;
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsOldCorrespondence_numRows += Repeat1__numRows;

// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsOldCorrespondence_total = rsOldCorrespondence.RecordCount;

// set the number of rows displayed on this page
if (rsOldCorrespondence_numRows < 0) {            // if repeat region set to all records
  rsOldCorrespondence_numRows = rsOldCorrespondence_total;
} else if (rsOldCorrespondence_numRows == 0) {    // if no repeat regions
  rsOldCorrespondence_numRows = 1;
}

// set the first and last displayed record
var rsOldCorrespondence_first = 1;
var rsOldCorrespondence_last  = rsOldCorrespondence_first + rsOldCorrespondence_numRows - 1;

// if we have the correct record count, check the other stats
if (rsOldCorrespondence_total != -1) {
  rsOldCorrespondence_numRows = Math.min(rsOldCorrespondence_numRows, rsOldCorrespondence_total);
  rsOldCorrespondence_first   = Math.min(rsOldCorrespondence_first, rsOldCorrespondence_total);
  rsOldCorrespondence_last    = Math.min(rsOldCorrespondence_last, rsOldCorrespondence_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsOldCorrespondence_total == -1) {

  // count the total records by iterating through the recordset
  for (rsOldCorrespondence_total=0; !rsOldCorrespondence.EOF; rsOldCorrespondence.MoveNext()) {
    rsOldCorrespondence_total++;
  }

  // reset the cursor to the beginning
  if (rsOldCorrespondence.CursorType > 0) {
    if (!rsOldCorrespondence.BOF) rsOldCorrespondence.MoveFirst();
  } else {
    rsOldCorrespondence.Requery();
  }

  // set the number of rows displayed on this page
  if (rsOldCorrespondence_numRows < 0 || rsOldCorrespondence_numRows > rsOldCorrespondence_total) {
    rsOldCorrespondence_numRows = rsOldCorrespondence_total;
  }

  // set the first and last displayed record
  rsOldCorrespondence_last  = Math.min(rsOldCorrespondence_first + rsOldCorrespondence_numRows - 1, rsOldCorrespondence_total);
  rsOldCorrespondence_first = Math.min(rsOldCorrespondence_first, rsOldCorrespondence_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsOldCorrespondence;
var MM_rsCount   = rsOldCorrespondence_total;
var MM_size      = rsOldCorrespondence_numRows;
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
rsOldCorrespondence_first = MM_offset + 1;
rsOldCorrespondence_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsOldCorrespondence_first = Math.min(rsOldCorrespondence_first, MM_rsCount);
  rsOldCorrespondence_last  = Math.min(rsOldCorrespondence_last, MM_rsCount);
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
	<title>Correspondence</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=550,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>	
</head>
<body>
<h5>DataSET Letters Sent</h5>
<table cellspacing="1">
	<tr>
		<td colspan="4">Displaying&nbsp;<%=rsOldCorrespondence_total%>&nbsp;Records.</td>
	</tr>	
</table>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
    <tr> 
		<th class="headrow" align="left" width="170">Letter Name</th>	
		<th class="headrow" align="left">Letter Type</th>
		<th class="headrow" align="left" width="150">Sender</th>	  
		<th class="headrow" align="left">Date Created</th>	  	  
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsOldCorrespondence.EOF)) { 
%>
    <tr> 
		<td nowrap><%=(rsOldCorrespondence.Fields.Item("chvLetter_Name").Value)%>&nbsp;</td>		
		<td nowrap><%=(rsOldCorrespondence.Fields.Item("chvLetter_type").Value)%>&nbsp;</td>		
		<td nowrap><%=(rsOldCorrespondence.Fields.Item("chvSender").Value)%>&nbsp;</td>
		<td nowraps><%=FilterDate(rsOldCorrespondence.Fields.Item("dtsSend_date").Value)%>&nbsp;</td>		
    </tr>
<%
	Repeat1__index++;
	rsOldCorrespondence.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Close Window" class="btnstyle" onClick="window.close();"></td>
	</tr>
</table>
</body>
</html>
<%
rsOldCorrespondence.Close();
%>
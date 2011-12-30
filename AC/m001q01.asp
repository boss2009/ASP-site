<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsClient__chvFilter = "dtsRefral_date >= '01/01/1900'";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
	rsClient__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Adult_Client(1,0,'"+ rsClient__chvFilter.replace(/'/g, "''") + "')}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
var rsClient_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsClient_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsClient_total = rsClient.RecordCount;

// set the number of rows displayed on this page
if (rsClient_numRows < 0) {            // if repeat region set to all records
  rsClient_numRows = rsClient_total;
} else if (rsClient_numRows == 0) {    // if no repeat regions
  rsClient_numRows = 1;
}

// set the first and last displayed record
var rsClient_first = 1;
var rsClient_last  = rsClient_first + rsClient_numRows - 1;

// if we have the correct record count, check the other stats
if (rsClient_total != -1) {
  rsClient_numRows = Math.min(rsClient_numRows, rsClient_total);
  rsClient_first   = Math.min(rsClient_first, rsClient_total);
  rsClient_last    = Math.min(rsClient_last, rsClient_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsClient_total == -1) {

  // count the total records by iterating through the recordset
  for (rsClient_total=0; !rsClient.EOF; rsClient.MoveNext()) {
    rsClient_total++;
  }

  // reset the cursor to the beginning
  if (rsClient.CursorType > 0) {
    if (!rsClient.BOF) rsClient.MoveFirst();
  } else {
    rsClient.Requery();
  }

  // set the number of rows displayed on this page
  if (rsClient_numRows < 0 || rsClient_numRows > rsClient_total) {
    rsClient_numRows = rsClient_total;
  }

  // set the first and last displayed record
  rsClient_last  = Math.min(rsClient_first + rsClient_numRows - 1, rsClient_total);
  rsClient_first = Math.min(rsClient_first, rsClient_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsClient;
var MM_rsCount   = rsClient_total;
var MM_size      = rsClient_numRows;
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
rsClient_first = MM_offset + 1;
rsClient_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsClient_first = Math.min(rsClient_first, MM_rsCount);
  rsClient_last  = Math.min(rsClient_last, MM_rsCount);
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
	<title>Client - Browse</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	
	function JumpRecord(){
		if (document.frmq01.JumpToRecord.value=="") {
			alert("Enter Record Number.");
		}
		if (!IsID(document.frmq01.JumpToRecord.value)) {
			alert("Invalid Record Number.");
		} else {
			window.location.href="..<%Response.Write(Request.ServerVariables("URL") + "?" + MM_keepMove + MM_moveParam + "=")%>"+String(document.frmq01.JumpToRecord.value-1);
		}
	}				
	</Script>
</head>
<body>
<form name="frmq01">
<h3>Client - Browse</h3>
<table cellspacing="1">
    <tr> 
		<td align="left" width="900">Displaying Records <b><%=(rsClient_first)%></b> to <b><%=(rsClient_last)%></b> of <b><%=(rsClient_total)%></b></td>
		<td nowrap><a href="javascript: openWindow('m001a0101a.asp','wQA01');">Add Client</a></td>		
    </tr>
</table>
<div class="BrowsePanel" style="width: 100%; height: 339px"> 
<table cellpadding="2" cellspacing="1">
      <tr> 
        <th nowrap class="headrow" align="left" width="180">Name</th>
        <th nowrap class="headrow" align="left">&nbsp;</th>
        <th nowrap class="headrow" align="center">SIN</th>
        <th nowrap class="headrow" align="left">Status</th>
        <th nowrap class="headrow" align="left">Case Manager</th>
        <th nowrap class="headrow" align="left">Disability</th>
<!--    
		<th nowrap class="headrow" align="left">Referral Date</th>
        <th nowrap class="headrow" align="left">Re-referral Date</th>
-->
        <% 
	if ( Session("MM_UserAuthorization") >= 6 ) { 
%>
        <th nowrap class="headrow" align="left">&nbsp;</th>
        <% 
	} 
%>
      </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsClient.EOF)) { 
%>
      <tr> 
        <td nowrap><a href="javascript: openWindow('m001FS3.asp?intAdult_id=<%=(rsClient.Fields.Item("intAdult_Id").Value)%>','wQE01');"><%=(rsClient.Fields.Item("chvLst_Name").Value)%>, 
          <%=(rsClient.Fields.Item("chvFst_Name").Value)%></a></td>
        <td nowrap><a href="javascript: openWindow('m001q01j.asp?intAdult_id=<%=rsClient.Fields.Item("intAdult_Id").Value%>','Summary');"><img src="../i/summary.gif" ALT="<%=(rsClient.Fields.Item("chvLst_Name").Value)%>, <%=(rsClient.Fields.Item("chvFst_Name").Value)%> summary info"></a></td>
        <td nowrap><%=FormatSIN(rsClient.Fields.Item("chrSIN_no").Value)%>&nbsp;</td>
        <td nowrap><%=(rsClient.Fields.Item("chvStatus").Value)%>&nbsp;</td>
        <td nowrap><%=(rsClient.Fields.Item("chvCaseManager").Value)%>&nbsp;</td>
        <td nowrap align="center"><%=(rsClient.Fields.Item("chvDisability").Value)%>&nbsp;</td>
<!--
        <td nowrap align="center"><%=FilterDate(rsClient.Fields.Item("dtsRefral_date").Value)%>&nbsp;</td>
        <td nowrap align="center"><%=FilterDate(rsClient.Fields.Item("dtsRe_refral_date").Value)%>&nbsp;</td>
-->
<% 
	if (Session("MM_UserAuthorization") >= 6 ) { 
%>
        <td nowrap><a href="javascript: openWindow('m001x01.asp?intAdult_id=<%=(rsClient.Fields.Item("intAdult_Id").Value)%>','wQE01');"><img src="../i/remove.gif" ALT="Remove <%=(rsClient.Fields.Item("chvLst_Name").Value)%>, <%=(rsClient.Fields.Item("chvFst_Name").Value)%>"></a></td>
<%
	} 
%>
      </tr>
<%
	Repeat1__index++;
	rsClient.MoveNext();
}
%>
</table>
</div>
</form>
</body>
</html>
<%
rsClient.Close();
%>
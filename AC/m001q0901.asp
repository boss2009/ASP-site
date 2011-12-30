<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
if (String(Request.QueryString("Remove"))=="True") {
	var rsRemoveLetter = Server.CreateObject("ADODB.Recordset");
	rsRemoveLetter.ActiveConnection = MM_cnnASP02_STRING;
	rsRemoveLetter.Source = "{call dbo.cp_delete_crspltr_custom("+ Request.QueryString("intLetter_id") + ",0)}";
	rsRemoveLetter.CursorType = 0;
	rsRemoveLetter.CursorLocation = 2;
	rsRemoveLetter.LockType = 3;
	rsRemoveLetter.Open();	
	Response.Redirect("m001q0901.asp?intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsOldCorrespondence = Server.CreateObject("ADODB.Recordset");
rsOldCorrespondence.ActiveConnection = MM_cnnASP02_STRING;
rsOldCorrespondence.Source = "{call dbo.cp_Clnt_Crsp("+ Request.QueryString("intAdult_id") + ")}";
rsOldCorrespondence.CursorType = 0;
rsOldCorrespondence.CursorLocation = 2;
rsOldCorrespondence.LockType = 3;
rsOldCorrespondence.Open();

var count = 0;

while (!rsOldCorrespondence.EOF) {
	count++;
	rsOldCorrespondence.MoveNext();
}

var rsCorrespondence = Server.CreateObject("ADODB.Recordset");
rsCorrespondence.ActiveConnection = MM_cnnASP02_STRING;
rsCorrespondence.Source = "{call dbo.cp_get_ac_crsp_hstry2("+ Request.QueryString("intAdult_id") + ",0)}";
rsCorrespondence.CursorType = 0;
rsCorrespondence.CursorLocation = 2;
rsCorrespondence.LockType = 3;
rsCorrespondence.Open();
var rsCorrespondence_numRows = 0;

var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsCorrespondence_numRows += Repeat1__numRows;

// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsCorrespondence_total = rsCorrespondence.RecordCount;

// set the number of rows displayed on this page
if (rsCorrespondence_numRows < 0) {            // if repeat region set to all records
  rsCorrespondence_numRows = rsCorrespondence_total;
} else if (rsCorrespondence_numRows == 0) {    // if no repeat regions
  rsCorrespondence_numRows = 1;
}

// set the first and last displayed record
var rsCorrespondence_first = 1;
var rsCorrespondence_last  = rsCorrespondence_first + rsCorrespondence_numRows - 1;

// if we have the correct record count, check the other stats
if (rsCorrespondence_total != -1) {
  rsCorrespondence_numRows = Math.min(rsCorrespondence_numRows, rsCorrespondence_total);
  rsCorrespondence_first   = Math.min(rsCorrespondence_first, rsCorrespondence_total);
  rsCorrespondence_last    = Math.min(rsCorrespondence_last, rsCorrespondence_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsCorrespondence_total == -1) {

  // count the total records by iterating through the recordset
  for (rsCorrespondence_total=0; !rsCorrespondence.EOF; rsCorrespondence.MoveNext()) {
    rsCorrespondence_total++;
  }

  // reset the cursor to the beginning
  if (rsCorrespondence.CursorType > 0) {
    if (!rsCorrespondence.BOF) rsCorrespondence.MoveFirst();
  } else {
    rsCorrespondence.Requery();
  }

  // set the number of rows displayed on this page
  if (rsCorrespondence_numRows < 0 || rsCorrespondence_numRows > rsCorrespondence_total) {
    rsCorrespondence_numRows = rsCorrespondence_total;
  }

  // set the first and last displayed record
  rsCorrespondence_last  = Math.min(rsCorrespondence_first + rsCorrespondence_numRows - 1, rsCorrespondence_total);
  rsCorrespondence_first = Math.min(rsCorrespondence_first, rsCorrespondence_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsCorrespondence;
var MM_rsCount   = rsCorrespondence_total;
var MM_size      = rsCorrespondence_numRows;
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
rsCorrespondence_first = MM_offset + 1;
rsCorrespondence_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsCorrespondence_first = Math.min(rsCorrespondence_first, MM_rsCount);
  rsCorrespondence_last  = Math.min(rsCorrespondence_last, MM_rsCount);
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
<h5>Correspondence</h5>
<table cellspacing="1">
	<tr>
		<td colspan="4">Displaying Records <%=rsCorrespondence_first%> To <%=rsCorrespondence_last%> Of <%=rsCorrespondence_total%></td>
	</tr>	
</table>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable" width="100%">
    <tr> 
		<th nowrap class="headrow" align="left">Letter Name</th>	
		<th nowrap class="headrow" align="left">Letter Type</th>
		<th nowrap class="headrow" align="left">Sender</th>	  
		<th nowrap class="headrow" align="left">Method</th>	  		
		<th nowrap class="headrow" align="left">Date Created</th>	  	  
		<th nowrap class="headrow">&nbsp;</th>		
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsCorrespondence.EOF)) { 
%>
    <tr>
<%
	switch (String(rsCorrespondence.Fields.Item("insTemplate_id").Value)) {
		case "0":
			if (Trim(rsCorrespondence.Fields.Item("chvRx_Type").Value)=="Custom Letter") {
%>		
		<td><a href="m001e0904.asp?intLetter_id=<%=(rsCorrespondence.Fields.Item("intLetter_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>"><%=(rsCorrespondence.Fields.Item("chvLetter_Name").Value)%></a></td>
<%
			} else {
%>
		<td><%=(rsCorrespondence.Fields.Item("chvLetter_Name").Value)%></td>
<%
			}
		break;
		//CSG Accept
		case "861":
%>
		<td><a href="m001e0902.asp?intLetter_id=<%=(rsCorrespondence.Fields.Item("intLetter_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>&insTemplate_id=<%=rsCorrespondence.Fields.Item("insTemplate_id").Value%>&intBuyout_req_id=<%=rsCorrespondence.Fields.Item("intBuyout_Req_id").Value%>"><%=(rsCorrespondence.Fields.Item("chvLetter_Name").Value)%></a></td>
<%		
		break;
		//Loan Accept
		case "860":
%>
		<td><a href="m001e0903.asp?intLetter_id=<%=(rsCorrespondence.Fields.Item("intLetter_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>&insTemplate_id=<%=rsCorrespondence.Fields.Item("insTemplate_id").Value%>&intLoan_Req_id=<%=rsCorrespondence.Fields.Item("intLoan_Req_id").Value%>"><%=(rsCorrespondence.Fields.Item("chvLetter_Name").Value)%></a></td>
<%
		break;
		//Client letters
		default:
%>
		<td><a href="m001e0901.asp?intLetter_id=<%=(rsCorrespondence.Fields.Item("intLetter_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>&insTemplate_id=<%=rsCorrespondence.Fields.Item("insTemplate_id").Value%>"><%=(rsCorrespondence.Fields.Item("chvLetter_Name").Value)%></a></td>
<%
		break;
	}	
%>
		<td nowrap><%=(rsCorrespondence.Fields.Item("chvRx_Type").Value)%>&nbsp;</td>		
		<td nowrap><%=(rsCorrespondence.Fields.Item("chvSender").Value)%>&nbsp;</td>
		<td nowrap><%=(rsCorrespondence.Fields.Item("chvSend_Method").Value)%>&nbsp;</td>		
		<td nowrap><%=FilterDate(rsCorrespondence.Fields.Item("dtsSend_Date").Value)%>&nbsp;</td>		
		<td nowrap><a href="m001q0901.asp?intLetter_id=<%=(rsCorrespondence.Fields.Item("intLetter_id").Value)%>&intAdult_id=<%=Request.QueryString("intAdult_id")%>&Remove=True"><img src="../i/Remove.gif" ALT="Remove <%=(rsCorrespondence.Fields.Item("chvLetter_Name").Value)%> Sent on <%=(rsCorrespondence.Fields.Item("dtsSend_Date").Value)%>"></a></td>		
    </tr>
<%
	Repeat1__index++;
	rsCorrespondence.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td width="150"><a href="javascript: openWindow('m001a0901.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>','winAdd');">Add Correspondence</a></td>
<%
if (count > 0) {
%>
		<td><a href="javascript: openWindow('m001q0902.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>','winAdd');">View DataSET Letters</a></td>		
<%
}
%>
	</tr>
</table>
</body>
</html>
<%
rsCorrespondence.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsPILATStudent__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsPILATStudent__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}

var rsPILATStudent__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsPILATStudent__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsPILATStudent__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsPILATStudent__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsPILATStudent = Server.CreateObject("ADODB.Recordset");
rsPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
rsPILATStudent.Source = "{call dbo.cp_pilat_student(0,'','','','','',0,0,0,0,'',0,0,0,"+rsPILATStudent__inspSrtBy+","+rsPILATStudent__inspSrtOrd+",'"+rsPILATStudent__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
rsPILATStudent.CursorType = 0;
rsPILATStudent.CursorLocation = 2;
rsPILATStudent.LockType = 3;
rsPILATStudent.Open();
var rsPILATStudent_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsPILATStudent_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsPILATStudent_total = rsPILATStudent.RecordCount;

// set the number of rows displayed on this page
if (rsPILATStudent_numRows < 0) {            // if repeat region set to all records
  rsPILATStudent_numRows = rsPILATStudent_total;
} else if (rsPILATStudent_numRows == 0) {    // if no repeat regions
  rsPILATStudent_numRows = 1;
}

// set the first and last displayed record
var rsPILATStudent_first = 1;
var rsPILATStudent_last  = rsPILATStudent_first + rsPILATStudent_numRows - 1;

// if we have the correct record count, check the other stats
if (rsPILATStudent_total != -1) {
  rsPILATStudent_numRows = Math.min(rsPILATStudent_numRows, rsPILATStudent_total);
  rsPILATStudent_first   = Math.min(rsPILATStudent_first, rsPILATStudent_total);
  rsPILATStudent_last    = Math.min(rsPILATStudent_last, rsPILATStudent_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsPILATStudent_total == -1) {

  // count the total records by iterating through the recordset
  for (rsPILATStudent_total=0; !rsPILATStudent.EOF; rsPILATStudent.MoveNext()) {
    rsPILATStudent_total++;
  }

  // reset the cursor to the beginning
  if (rsPILATStudent.CursorType > 0) {
    if (!rsPILATStudent.BOF) rsPILATStudent.MoveFirst();
  } else {
    rsPILATStudent.Requery();
  }

  // set the number of rows displayed on this page
  if (rsPILATStudent_numRows < 0 || rsPILATStudent_numRows > rsPILATStudent_total) {
    rsPILATStudent_numRows = rsPILATStudent_total;
  }

  // set the first and last displayed record
  rsPILATStudent_last  = Math.min(rsPILATStudent_first + rsPILATStudent_numRows - 1, rsPILATStudent_total);
  rsPILATStudent_first = Math.min(rsPILATStudent_first, rsPILATStudent_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsPILATStudent;
var MM_rsCount   = rsPILATStudent_total;
var MM_size      = rsPILATStudent_numRows;
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
rsPILATStudent_first = MM_offset + 1;
rsPILATStudent_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsPILATStudent_first = Math.min(rsPILATStudent_first, MM_rsCount);
  rsPILATStudent_last  = Math.min(rsPILATStudent_last, MM_rsCount);
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
	<title>Temp Student - Browse</title>
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
<h3>Temp Student - Browse</h3>
<table cellspacing="1">
    <tr> 
		<td align="left" width="320">Displaying Records <b><%=(rsPILATStudent_first)%></b> to <b><%=(rsPILATStudent_last)%></b> of <b><%=(rsPILATStudent_total)%></b></td>
		<td nowrap><a href="javascript: openWindow('m022a0101.asp','wQA22');">Add PILAT Student</a></td>
    </tr>
</table>
<div class="BrowsePanel" style="height: 320px;"> 
<table cellpadding="2" cellspacing="1">
	<tr> 
        <th nowrap class="headrow" align="left" width="200">Student Name</th>
        <th nowrap class="headrow" align="center" width="60">SIN</th>
        <th nowrap class="headrow" align="center">Primary Disability</th>
        <th nowrap class="headrow" align="center" width="70">Status</th>
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsPILATStudent.EOF)) { 
%>
    <tr> 
        <td nowrap><a href="javascript: openWindow('m022FS3.asp?intPStdnt_id=<%=(rsPILATStudent.Fields.Item("intPStdnt_id").Value)%>','wQE01');"><%=(rsPILATStudent.Fields.Item("chvLst_Name").Value)%>,&nbsp;<%=(rsPILATStudent.Fields.Item("chvFst_Name").Value)%></a></td>
        <td nowrap><%=FormatSIN(rsPILATStudent.Fields.Item("chrSIN_no").Value)%>&nbsp;</td>
        <td align="center" nowrap><%=(rsPILATStudent.Fields.Item("chvDisability").Value)%>&nbsp;</td>
        <td nowrap><%=(rsPILATStudent.Fields.Item("chvStdnt_Status").Value)%>&nbsp;</td>
    </tr>
<%
	Repeat1__index++;
	rsPILATStudent.MoveNext();
}
%>
</table>
</div>
</form>
</body>
</html>
<%
rsPILATStudent.Close();
%>

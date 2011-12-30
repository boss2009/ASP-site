<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsCompany__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsCompany__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}
var rsCompany__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsCompany__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsCompany__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsCompany__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsCompany = Server.CreateObject("ADODB.Recordset");
rsCompany.ActiveConnection = MM_cnnASP02_STRING;
rsCompany.Source = "{call dbo.cp_company2(0,'',0,0,0,0,0,"+rsCompany__inspSrtBy.replace(/'/g, "''")+","+rsCompany__inspSrtOrd.replace(/'/g, "''")+",'"+rsCompany__chvFilter.replace(/'/g, "''")+"',0,'Q',0)}";
rsCompany.CursorType = 0;
rsCompany.CursorLocation = 2;
rsCompany.LockType = 3;
rsCompany.Open();
var rsCompany_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsCompany_numRows += Repeat1__numRows;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsCompany_total = rsCompany.RecordCount;

// set the number of rows displayed on this page
if (rsCompany_numRows < 0) {            // if repeat region set to all records
  rsCompany_numRows = rsCompany_total;
} else if (rsCompany_numRows == 0) {    // if no repeat regions
  rsCompany_numRows = 1;
}

// set the first and last displayed record
var rsCompany_first = 1;
var rsCompany_last  = rsCompany_first + rsCompany_numRows - 1;

// if we have the correct record count, check the other stats
if (rsCompany_total != -1) {
  rsCompany_numRows = Math.min(rsCompany_numRows, rsCompany_total);
  rsCompany_first   = Math.min(rsCompany_first, rsCompany_total);
  rsCompany_last    = Math.min(rsCompany_last, rsCompany_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsCompany_total == -1) {

  // count the total records by iterating through the recordset
  for (rsCompany_total=0; !rsCompany.EOF; rsCompany.MoveNext()) {
    rsCompany_total++;
  }

  // reset the cursor to the beginning
  if (rsCompany.CursorType > 0) {
    if (!rsCompany.BOF) rsCompany.MoveFirst();
  } else {
    rsCompany.Requery();
  }

  // set the number of rows displayed on this page
  if (rsCompany_numRows < 0 || rsCompany_numRows > rsCompany_total) {
    rsCompany_numRows = rsCompany_total;
  }

  // set the first and last displayed record
  rsCompany_last  = Math.min(rsCompany_first + rsCompany_numRows - 1, rsCompany_total);
  rsCompany_first = Math.min(rsCompany_first, rsCompany_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsCompany;
var MM_rsCount   = rsCompany_total;
var MM_size      = rsCompany_numRows;
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
rsCompany_first = MM_offset + 1;
rsCompany_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsCompany_first = Math.min(rsCompany_first, MM_rsCount);
  rsCompany_last  = Math.min(rsCompany_last, MM_rsCount);
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
	<title>Organizations - Browse</title>
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
<h3>Organizations - Browse</h3>
<table cellspacing="1">
    <tr> 
		<td align="left" width="900">Displaying Records <b><%=(rsCompany_first)%></b> to <b><%=(rsCompany_last)%></b> of <b><%=(rsCompany_total)%></b></td>
		<td nowrap><a href="javascript: openWindow('m006a0101.asp','wQA06');">Add Organization</a></td>    
	</tr>
</table>
<div class="BrowsePanel" style="width: 100%; height: 320px"> 
<table cellpadding="2" cellspacing="1">
	<tr> 
        <th nowrap class="headrow" align="left" width="300">Organization Name</th>
        <th nowrap class="headrow" align="left">Type</th>
        <th nowrap class="headrow" align="left">Address</th>
        <th nowrap class="headrow" align="center">City</th>
        <th nowrap class="headrow" align="center">Province/State</th>
        <th nowrap class="headrow" align="center">Country</th>
        <th nowrap class="headrow" align="left">Phone Number</th>
    </tr>
<% 
while ((Repeat1__numRows-- != 0) && (!rsCompany.EOF)) { 
%>
      <tr> 
        <td nowrap><a href="javascript: openWindow('m006FS3.asp?intCompany_id=<%=(rsCompany.Fields.Item("intCompany_id").Value)%>','wQE01');"><%=(rsCompany.Fields.Item("chvCompany_Name").Value)%></a></td>
        <td nowrap valign="top"><%=(rsCompany.Fields.Item("chvWork_type_desc").Value)%>&nbsp;</td>
        <td nowrap valign="top"><%=(rsCompany.Fields.Item("chvAddress").Value)%>&nbsp;</td>
        <td nowrap valign="top" align="center"><%=(rsCompany.Fields.Item("chvCity").Value)%>&nbsp;</td>
        <td nowrap valign="top" align="center"><%=(rsCompany.Fields.Item("chrprvst_abbv").Value)%>&nbsp;</td>
        <td nowrap valign="top" align="center"><%=(rsCompany.Fields.Item("chvcntry_name").Value)%>&nbsp;</td>
        <td nowrap valign="top"><%=FormatPhoneNumber(rsCompany.Fields.Item("chvPhone_Type_1").Value,rsCompany.Fields.Item("chvPhone1_Arcd").Value,rsCompany.Fields.Item("chvPhone1_Num").Value,rsCompany.Fields.Item("chvPhone1_Ext").Value,rsCompany.Fields.Item("chvPhone_Type_2").Value,rsCompany.Fields.Item("chvPhone2_Arcd").Value,rsCompany.Fields.Item("chvPhone2_Num").Value,rsCompany.Fields.Item("chvPhone2_Ext").Value,rsCompany.Fields.Item("chvPhone_Type_3").Value,rsCompany.Fields.Item("chvPhone3_Arcd").Value,rsCompany.Fields.Item("chvPhone3_Num").Value,rsCompany.Fields.Item("chvPhone3_Ext").Value,rsCompany.Fields.Item("chvPhone3_Ext").Value)%>&nbsp;</td>
      </tr>
<%
	Repeat1__index++;
	rsCompany.MoveNext();
}
%>
</table>
</div>
</form>
</body>
</html>
<%
rsCompany.Close();
%>
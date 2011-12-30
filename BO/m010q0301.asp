<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var gst = 0;
var pst = 0;
var shipping = 0;

var rsGST = Server.CreateObject("ADODB.Recordset");
rsGST.ActiveConnection = MM_cnnASP02_STRING;
rsGST.Source = "{call dbo.cp_charge_rate(1,'',0,0.0,1,'Q',0)}";
rsGST.CursorType = 0;
rsGST.CursorLocation = 2;
rsGST.LockType = 3;
rsGST.Open();
if (!rsGST.EOF) gst = Number(rsGST.Fields.Item("fltPercentage").Value);
rsGST.Close();

var rsPST = Server.CreateObject("ADODB.Recordset");
rsPST.ActiveConnection = MM_cnnASP02_STRING;
rsPST.Source = "{call dbo.cp_charge_rate(2,'',0,0.0,1,'Q',0)}";
rsPST.CursorType = 0;
rsPST.CursorLocation = 2;
rsPST.LockType = 3;
rsPST.Open();
if (!rsPST.EOF) pst = Number(rsPST.Fields.Item("fltPercentage").Value);
rsPST.Close();

var rsShipping = Server.CreateObject("ADODB.Recordset");
rsShipping.ActiveConnection = MM_cnnASP02_STRING;
rsShipping.Source = "{call dbo.cp_charge_rate(3,'',0,0.0,1,'Q',0)}";
rsShipping.CursorType = 0;
rsShipping.CursorLocation = 2;
rsShipping.LockType = 3;
rsShipping.Open();
if (!rsShipping.EOF) shipping = Number(rsShipping.Fields.Item("fltPercentage").Value);
rsShipping.Close();

var rsInventorySold = Server.CreateObject("ADODB.Recordset");
rsInventorySold.ActiveConnection = MM_cnnASP02_STRING;
rsInventorySold.Source = "{call dbo.cp_buyout_eqp_sold(0,"+Request.QueryString("intBuyout_req_id")+",0,0.0,'',0,0,'',0,'Q',0)}";
rsInventorySold.CursorType = 0;
rsInventorySold.CursorLocation = 2;
rsInventorySold.LockType = 3;
rsInventorySold.Open();
var rsInventorySold_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsInventorySold_numRows += Repeat1__numRows;
%>
<%
// set the record count
var rsInventorySold_total = rsInventorySold.RecordCount;

// set the number of rows displayed on this page
if (rsInventorySold_numRows < 0) {            // if repeat region set to all records
  rsInventorySold_numRows = rsInventorySold_total;
} else if (rsInventorySold_numRows == 0) {    // if no repeat regions
  rsInventorySold_numRows = 1;
}

// set the first and last displayed record
var rsInventorySold_first = 1;
var rsInventorySold_last  = rsInventorySold_first + rsInventorySold_numRows - 1;

// if we have the correct record count, check the other stats
if (rsInventorySold_total != -1) {
  rsInventorySold_numRows = Math.min(rsInventorySold_numRows, rsInventorySold_total);
  rsInventorySold_first   = Math.min(rsInventorySold_first, rsInventorySold_total);
  rsInventorySold_last    = Math.min(rsInventorySold_last, rsInventorySold_total);
}

// *** Recordset Stats: if we don't know the record count, manually count them

if (rsInventorySold_total == -1) {

  // count the total records by iterating through the recordset
  for (rsInventorySold_total=0; !rsInventorySold.EOF; rsInventorySold.MoveNext()) {
    rsInventorySold_total++;
  }

  // reset the cursor to the beginning
  if (rsInventorySold.CursorType > 0) {
    if (!rsInventorySold.BOF) rsInventorySold.MoveFirst();
  } else {
    rsInventorySold.Requery();
  }

  // set the number of rows displayed on this page
  if (rsInventorySold_numRows < 0 || rsInventorySold_numRows > rsInventorySold_total) {
    rsInventorySold_numRows = rsInventorySold_total;
  }

  // set the first and last displayed record
  rsInventorySold_last  = Math.min(rsInventorySold_first + rsInventorySold_numRows - 1, rsInventorySold_total);
  rsInventorySold_first = Math.min(rsInventorySold_first, rsInventorySold_total);
}
var MM_paramName = "";

// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsInventorySold;
var MM_rsCount   = rsInventorySold_total;
var MM_size      = rsInventorySold_numRows;
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
rsInventorySold_first = MM_offset + 1;
rsInventorySold_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsInventorySold_first = Math.min(rsInventorySold_first, MM_rsCount);
  rsInventorySold_last  = Math.min(rsInventorySold_last, MM_rsCount);
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
	<title>Equipment Sold</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>
</head>
<body>
<h5>Equipment Sold</h5>
<table cellspacing="1">
    <tr> 
		<td nowrap width="450"><a href="javascript: openWindow('m010a0301.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>','w010A02');">Add Equipment Sold</a></td>			
    	<td nowrap align="left">Displaying <b><%=(rsInventorySold_total)%></b> Records.</td>
	</tr>
</table>
<hr>
<div class="BrowsePanel" style="width: 100%; height: 220px"> 
<table cellpadding="2" cellspacing="1">
    <tr> 
		<th nowrap class="headrow" align="left">Inventory Name</th>
		<th nowrap class="headrow" align="left">Inventory ID</th>
		<th nowrap class="headrow" align="left">Status</th>
		<th nowrap class="headrow" align="left">Date Processed</th>
		<th nowrap class="headrow" align="left">Sold Price</th>
		<th nowrap class="headrow" align="left">Equipment Cost</th>
		<th nowrap class="headrow" align="left">Serial Number</th>
		<th nowrap class="headrow" align="left">Model Number</th>		
		<th nowrap class="headrow" align="left">PR Number</th>
		<th nowrap class="headrow" align="left">Vendor</th>
		<th nowrap class="headrow" align="left">Date Returned</th>
		<th nowrap class="headrow" align="left">Returned By</th>
		<th nowrap class="headrow" align="left">Comments</th>
    </tr>
<% 
var total_sold_price = 0;
var total_cost = 0;
var tax = 0;
var total_shipping = 0;
while (!rsInventorySold.EOF) { 
	if (!(rsInventorySold.Fields.Item("insEquip_Class_id").Value==null)) {		
		var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
		rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
		rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW(" + rsInventorySold.Fields.Item("insEquip_Class_id").Value + ",'C',1)}";	
		rsConcreteClass.CursorType = 0;
		rsConcreteClass.CursorLocation = 2;
		rsConcreteClass.LockType = 3;
		rsConcreteClass.Open();	
		switch (String(rsConcreteClass.Fields.Item("chvSbjTotax").Value)) {
			//pst
			case "1":
				tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * (pst/100));
			break;
			//gst
			case "2":
				tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * (gst/100));		
			break;
			//both
			case "3":
				tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * ((gst+pst)/100));		
			break;
		}
	}
	total_cost += rsInventorySold.Fields.Item("fltList_Unit_Cost").Value;
	total_sold_price += rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value;		
%>
    <tr> 
		<td nowrap align="left"><a href="m010e0301.asp?intBO_Eqp_Sold_id=<%=rsInventorySold.Fields.Item("intBO_Eqp_Sold_id").Value%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>"><%=(rsInventorySold.Fields.Item("chvInventory_Name").Value)%></a>&nbsp;</td>
		<td nowrap align="center"><%=ZeroPadFormat(rsInventorySold.Fields.Item("intEquip_set_id").Value,8)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventorySold.Fields.Item("chvEqp_Status").Value)%>&nbsp;</td>
		<td nowrap align="center"><%=FilterDate(rsInventorySold.Fields.Item("dtsDate_processed").Value)%>&nbsp;</td>
		<td nowrap align="right"><%=FormatCurrency(rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value)%>&nbsp;</td>		
		<td nowrap align="right"><%=FormatCurrency(rsInventorySold.Fields.Item("fltList_Unit_Cost").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventorySold.Fields.Item("chvSerial_Number").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventorySold.Fields.Item("chvModel_Number").Value)%>&nbsp;</td>
		<td nowrap align="center"><%=ZeroPadFormat(rsInventorySold.Fields.Item("intRequisition_no").Value,8)%>&nbsp;</td>						
		<td nowrap align="left"><%=(rsInventorySold.Fields.Item("chvVendor").Value)%>&nbsp;</td>		
		<td nowrap align="center"><%=FilterDate(rsInventorySold.Fields.Item("dtsDate_Returned").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventorySold.Fields.Item("chvRtned_by").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventorySold.Fields.Item("chvComments").Value)%>&nbsp;</td>
    </tr>
<%
	rsInventorySold.MoveNext();
}
total_shipping = total_sold_price * (shipping/100);
%>
  </table>
</div>
<div style="position: absolute; top: 310px">
<table cellpadding="0" cellspacing="1">
	<tr>
		<td width="350"><b>Total Equipment Cost:</b></td>
		<td align="right"><b><%=FormatCurrency(total_cost)%></b></td>
	</tr>
	<tr>
		<td width="350"><b>Total Sold Cost without taxes/shipping:</b></td>
		<td align="right"><b><%=FormatCurrency(total_sold_price)%></b></td>
	</tr>
	<tr>
		<td width="350"><b>Taxes:</b></td>
		<td align="right"><b><%=FormatCurrency(tax)%></b></td>
	</tr>
	<tr>
		<td width="350"><b>Shipping:</b></td>
		<td align="right"><b><%=FormatCurrency(total_shipping)%></b></td>
	</tr>	
	<tr>	
		<td width="350"><b>Total Buyout Cost with taxes/shipping:</b></td>
		<td align="right"><b><%=FormatCurrency(total_sold_price+tax+total_shipping)%></b></td>
	</tr>
</table>
</div>
</body>
</html>
<%
rsInventorySold.Close();
%>
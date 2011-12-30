<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsInventoryRequested = Server.CreateObject("ADODB.Recordset");
rsInventoryRequested.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryRequested.Source = "{call dbo.cp_eqp_requested(0,"+Request.QueryString("intLoan_req_id")+",0,0,0,'',0.0,0,0,'Q',0)}";
rsInventoryRequested.CursorType = 0;
rsInventoryRequested.CursorLocation = 2;
rsInventoryRequested.LockType = 3;
rsInventoryRequested.Open();

var count = 0;
while (!rsInventoryRequested.EOF) {
	count++;
	rsInventoryRequested.MoveNext();
}
%>
<html>
<head>
	<title>Backordered Item</title>
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
<h5>Backordered Item</h5>
<table cellspacing="1">
	<tr> 
		<td colspan="4" align="left">Displaying <b><%=count%></b> Records.</td>
	</tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
    <tr> 
		<th nowrap class="headrow" align="left" width="250">Class/Bundle Name</th>
		<th nowrap class="headrow" align="center">Type</th>		
		<th nowrap class="headrow" align="center">Quantity</th>
		<th nowrap class="headrow" align="left" width="250">Comments</th>
    </tr>
<% 
while ((!rsInventoryRequested.EOF)) { 
	if (rsInventoryRequested.Fields.Item("bitIs_BO").Value=="1") {
%>
	<tr> 
		<td valign="top" align="left" nowrap><a href="m008e0201.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intEqpReq_Id=<%=(rsInventoryRequested.Fields.Item("intEqpReq_Id").Value)%>"><%=((rsInventoryRequested.Fields.Item("bitIs_class").Value=="1")?rsInventoryRequested.Fields.Item("chvEqp_Class_Name").Value:rsInventoryRequested.Fields.Item("chvEqp_Bundle_Name").Value)%></a>&nbsp;</td>
		<td valign="top" align="left" nowrap><%=((rsInventoryRequested.Fields.Item("bitIs_class").Value=="1")?"Class":"Bundle")%></td>		
		<td valign="top" align="center" nowrap><%=(rsInventoryRequested.Fields.Item("insQuantity").Value)%>&nbsp;</td>
		<td valign="top" align="left"><%=(rsInventoryRequested.Fields.Item("chvComments").Value)%>&nbsp;</td>
	</tr>
<%
		if (rsInventoryRequested.Fields.Item("bitIs_class").Value == "0") {
			var rsBundleComponent = Server.CreateObject("ADODB.Recordset");
			rsBundleComponent.ActiveConnection = MM_cnnASP02_STRING;
			rsBundleComponent.Source = "{call dbo.cp_bundle_eqp_class("+rsInventoryRequested.Fields.Item("insClass_bundle_id").Value+",0,0,'Q',0)}";
			rsBundleComponent.CursorType = 0;
			rsBundleComponent.CursorLocation = 2;
			rsBundleComponent.LockType = 3;
			rsBundleComponent.Open();
			while (!rsBundleComponent.EOF) {
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsInventoryRequested_total = rsInventoryRequested.RecordCount;

// set the number of rows displayed on this page
if (rsInventoryRequested_numRows < 0) {            // if repeat region set to all records
  rsInventoryRequested_numRows = rsInventoryRequested_total;
} else if (rsInventoryRequested_numRows == 0) {    // if no repeat regions
  rsInventoryRequested_numRows = 1;
}

// set the first and last displayed record
var rsInventoryRequested_first = 1;
var rsInventoryRequested_last  = rsInventoryRequested_first + rsInventoryRequested_numRows - 1;

// if we have the correct record count, check the other stats
if (rsInventoryRequested_total != -1) {
  rsInventoryRequested_numRows = Math.min(rsInventoryRequested_numRows, rsInventoryRequested_total);
  rsInventoryRequested_first   = Math.min(rsInventoryRequested_first, rsInventoryRequested_total);
  rsInventoryRequested_last    = Math.min(rsInventoryRequested_last, rsInventoryRequested_total);
}
%>

<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsInventoryRequested_total == -1) {

  // count the total records by iterating through the recordset
  for (rsInventoryRequested_total=0; !rsInventoryRequested.EOF; rsInventoryRequested.MoveNext()) {
    rsInventoryRequested_total++;
  }

  // reset the cursor to the beginning
  if (rsInventoryRequested.CursorType > 0) {
    if (!rsInventoryRequested.BOF) rsInventoryRequested.MoveFirst();
  } else {
    rsInventoryRequested.Requery();
  }

  // set the number of rows displayed on this page
  if (rsInventoryRequested_numRows < 0 || rsInventoryRequested_numRows > rsInventoryRequested_total) {
    rsInventoryRequested_numRows = rsInventoryRequested_total;
  }

  // set the first and last displayed record
  rsInventoryRequested_last  = Math.min(rsInventoryRequested_first + rsInventoryRequested_numRows - 1, rsInventoryRequested_total);
  rsInventoryRequested_first = Math.min(rsInventoryRequested_first, rsInventoryRequested_total);
}
%>

	<tr>
		<td nowrap colspan="4" style="font-size: 7pt">&nbsp;&nbsp;-&nbsp;<%=rsBundleComponent.Fields.Item("chvEqCls_name").Value%></td>
	</tr>
<%
				rsBundleComponent.MoveNext();
			}
			rsBundleComponent.Close();
		}
	}
	rsInventoryRequested.MoveNext();
}
%>
</table>
<br><br><br>
<input type="button" value="Close" onclick="window.close();" class="btnstyle">
</body>
</html>
<%
rsInventoryRequested.Close();
%>
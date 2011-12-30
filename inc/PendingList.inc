<%@Language="JavaScript"%>
<%
/*
var Heather = 39

var Rachelle = 140

var Daniel = 136
var Tim = 135

var Vince = 139
var Ron =
var Garth = 87

var Gladys = 29
var Deanna = 25

var William = 129
var Douglas = 28
*/
switch (Request.Form("PendingItems")){
	<!-- Outstanding Buyout Requested -->
	case "1":
		var rsBuyout = Server.CreateObject("ADODB.Recordset");
		rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
		rsBuyout.Source = "{call dbo.cp_search_buyout_request(5,1,'insBuyout_Status_id = 1',0)}";
		rsBuyout.CursorType = 0;
		rsBuyout.CursorLocation = 2;
		rsBuyout.LockType = 3;
		rsBuyout.Open();
%>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr>
		<th nowrap class="headrow" align="left">Buyout ID</th>
	        <th nowrap class="headrow" align="left">Buyer Name</th>
	</tr>
<%
while (!rsBuyout.EOF) {
%>
	<tr>
		<td nowrap align="left"><a href="javascript: openWindow('m010FS3.asp?intBuyout_Req_id=<%=(rsBuyout.Fields.Item("intBuyout_Req_id").Value)%>','w010E01');"><%=ZeroPadFormat(rsBuyout.Fields.Item("intBuyout_Req_id").Value, 8)%></a></td>
        	<td nowrap align="left"><%=((rsBuyout.Fields.Item("insEq_user_type").Value==3)?rsBuyout.Fields.Item("chvLst_Name").Value+", "+rsBuyout.Fields.Item("chvFst_Name").Value:rsBuyout.Fields.Item("chvSchool_Name").Value)%>&nbsp;</td>
	</tr>
<%
	rsBuyout.MoveNext();
}
%>
</table>
<%
	break;
}
<!-- Outstanding Buyout Approved -->
// insBuyout_Status_id = 2

<!-- Outstanding Loan Requested -->
// insLoan_Status_id = 1

<!-- Outstanding Loan Approved -->
// insLoan_Status_id = 2

<!-- Outstanding Repair Jobs -->
// insRepair_Status = 1

<!-- Outstanding CIP -->
%>
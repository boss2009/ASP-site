<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup2(758,0,'0',0,'1',0)}";
rsCol.CursorType = 0;
rsCol.CursorLocation = 2;
rsCol.LockType = 3;
rsCol.Open();

var rsUserType = Server.CreateObject("ADODB.Recordset");
rsUserType.ActiveConnection = MM_cnnASP02_STRING;
rsUserType.Source = "{call dbo.cp_eq_user_type2(0,'',1,0,0,'Q',0)}";
rsUserType.CursorType = 0;
rsUserType.CursorLocation = 2;
rsUserType.LockType = 3;
rsUserType.Open();

var rsBuyoutStatus = Server.CreateObject("ADODB.Recordset");
rsBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutStatus.Source = "{call dbo.cp_buyout_status(0,'',0,'Q',0)}";
rsBuyoutStatus.CursorType = 0;
rsBuyoutStatus.CursorLocation = 2;
rsBuyoutStatus.LockType = 3;
rsBuyoutStatus.Open();
%>
<html>
<head>
	<title>Buyout Equipment Request To Do</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript">
	function Search() {
		var inspSrtBy = document.frm10s01.SortByColumn.value;
		var inspSrtOrd = document.frm10s01.OrderBy.value;
		var stgFilter = "" ;
		stgFilter = "insEq_user_type = " + document.frm10s01.UserType.value + " ";
		var l = "";
		var	m = document.frm10s01.BuyoutStatus.length;
		for (var ii = 0; ii < m; ii++) {
			if (document.frm10s01.BuyoutStatus[ii].selected) {
				if (l.length > 0) l += "," ;
				l += document.frm10s01.BuyoutStatus[ii].value ;	
			} 				  
		} 	
		stgFilter += " AND insBuyout_Status_id in (" + l + ") " ; 
		document.frm10s01.action = "m010q02.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;				
		document.frm10s01.submit();
	}

	if (window.focus) self.focus();		
	</script>
</head>
<body>
<form name="frm10s01" method="post" action="">
<h3>Buyout Equipment Request To Do</h3>
<i>Hold down [Ctrl] key to select multiple Loan Status</i>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td width="80">User Type</td>
		<td><select name="UserType" tabindex="1" style="width: 200px">
			<% 
			while (!rsUserType.EOF) { 
			%>
				<option value="<%=(rsUserType.Fields.Item("insEq_user_type").Value)%>"><%=(rsUserType.Fields.Item("chvEq_user_type").Value)%>
			<% 
				rsUserType.MoveNext();
			}
			%>
		</select></td>
    </tr>
	<tr>
		<td valign="top">Buyout Status</td>
		<td><select name="BuyoutStatus" tabindex="2" size="5" multiple style="width: 200px">
			<% 
			while (!rsBuyoutStatus.EOF) { 
			%>
				<option value="<%=rsBuyoutStatus("insbuyout_status_id")%>"><%=rsBuyoutStatus("chvBuyout_status")%>
			<%
				rsBuyoutStatus.MoveNext 
			}
			%>
		</select></td>		
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>
			Sort by:
			<select name="SortByColumn" tabindex="3">
			<% 
			while (!rsCol.EOF) {
			%>
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value=="5")?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvObjName").Value)%></option>
			<%
				rsCol.MoveNext();
			}
			%>
			</select>
			Order 
        	<select name="OrderBy" tabindex="4">
				<option value="0">Ascending</option>
				<option value="1">Descending</option>
			</select>
		</td>
	</tr>
	<tr>
		<td>
	        <input type="button" value="Search" onClick="Search();" tabindex="5" class="btnstyle">
			<input type="reset" value="Clear All" onClick="window.location.reload();" tabindex="6" class="btnstyle">
		</td>		
    </tr>
</table>
</form>
</body>
</html>
<%
rsCol.Close();
rsBuyoutStatus.Close();
%>
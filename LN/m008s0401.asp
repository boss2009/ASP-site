<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// retrieve sort columns
var rsCol = Server.CreateObject("ADODB.Recordset");
rsCol.ActiveConnection = MM_cnnASP02_STRING;
rsCol.Source = "{call dbo.cp_ASP_Lkup2(753,0,'',0,'',0)}";
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

var rsLoanStatus = Server.CreateObject("ADODB.Recordset");
rsLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
rsLoanStatus.Source = "{call dbo.cp_loan_status2(0,'',0,'Q',0)}";
rsLoanStatus.CursorType = 0;
rsLoanStatus.CursorLocation = 2;
rsLoanStatus.LockType = 3;
rsLoanStatus.Open();
%>
<html>
<head>
	<title>Loan Equipment Request To Do</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript" src="../js/m008Srh01.js"></script>
	<script language="JavaScript">
	
	if (window.focus) self.focus();

	function Search() {
		var inspSrtBy = document.frm08s01.SortByColumn.value;
		var inspSrtOrd = document.frm08s01.OrderBy.value;
		var stgFilter = "" ;
		stgFilter = ACfltr_08("184","","",document.frm08s01.UserType.value,"");

		var l = "";
		var	m = document.frm08s01.LoanStatus.length;
		for (var ii = 0; ii < m; ii++) {
			if (document.frm08s01.LoanStatus[ii].selected) {
				if (l.length > 0) l += "," ;
				l += document.frm08s01.LoanStatus[ii].value ;	
			} 				  
		} 	
		stgFilter += " AND insLoan_Status_id in (" + l + ") " ; 
		document.frm08s01.action = "m008q02.asp?inspSrtBy="+inspSrtBy+"&inspSrtOrd="+inspSrtOrd+"&chvFilter=" + stgFilter ;				
		document.frm08s01.submit();
	}
	</script>
</head>
<body>
<form name="frm08s01" method="post" action="">
<h3>Loan Equipment Request To Do</h3>
<i>Hold down [Ctrl] key to select multiple Loan Status</i>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td width="100">User Type</td>
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
		<td valign="top">Loan Status</td>
		<td valign="top"><select name="LoanStatus" tabindex="2" MULTIPLE size="5" style="width: 200px">
		<% 
		while (!rsLoanStatus.EOF) {
		%>
			<option value="<%=(rsLoanStatus.Fields.Item("intloan_status_id").Value)%>"><%=(rsLoanStatus.Fields.Item("chvname").Value)%></option>
		<%
			rsLoanStatus.MoveNext();
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
				<option value="<%=(rsCol.Fields.Item("insObjOrder").Value)%>" <%=((rsCol.Fields.Item("insObjOrder").Value==1)?"SELECTED":"")%>><%=(rsCol.Fields.Item("chvObjName").Value)%></option>
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
			<input type="reset" value="Clear All" tabindex="6" onClick="window.location.reload();" class="btnstyle">
		</td>		
	</tr>
</table>
</form>
</body>
</html>
<%
rsCol.Close();
rsLoanStatus.Close();
rsUserType.Close();
%>
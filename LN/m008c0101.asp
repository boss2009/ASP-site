<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.Form("MM_update")) == "true") {
	var LoanArray
	LoanArray = String(Request.Form("PriorityList")).split(":");
	if (LoanArray.length==1) {
		var rsRankLoan = Server.CreateObject("ADODB.Recordset");
		rsRankLoan.ActiveConnection = MM_cnnASP02_STRING;
		rsRankLoan.Source = "{call dbo.cp_update_loan_req_rank("+LoanArray+",0,0)}";
		rsRankLoan.CursorType = 0;
		rsRankLoan.CursorLocation = 2;
		rsRankLoan.LockType = 3;
		rsRankLoan.Open();	
	}
	if (LoanArray.length > 1) {
		var rsRankLoan = Server.CreateObject("ADODB.Recordset");
		rsRankLoan.ActiveConnection = MM_cnnASP02_STRING;
		rsRankLoan.Source = "{call dbo.cp_update_loan_req_rank("+LoanArray[0]+",0,0)}";
		rsRankLoan.CursorType = 0;
		rsRankLoan.CursorLocation = 2;
		rsRankLoan.LockType = 3;
		rsRankLoan.Open();	
		for (var i = 1; i < LoanArray.length; i ++) {
			rsRankLoan.Source = "{call dbo.cp_update_loan_req_rank("+LoanArray[i]+","+i+",0)}";
			rsRankLoan.Open();
		}		
	}
}

var rsLoans = Server.CreateObject("ADODB.Recordset");
rsLoans.ActiveConnection = MM_cnnASP02_STRING;
rsLoans.Source = "{call dbo.cp_loan_request2(0,0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,4,0,'',2,'Q',0)}";
rsLoans.CursorType = 0;
rsLoans.CursorLocation = 2;
rsLoans.LockType = 3;
rsLoans.Open();
%>
<html>
<head>
	<title>Work Priority</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">	
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=20,top=20,status=1");
		return ;
	}	   
	
	function Save(){
		if (document.frm0101.Count.value > 0) {
			document.frm0101.PriorityList.value = document.frm0101.ApprovedLoans.options[0].value;
			for (var i = 1; i < document.frm0101.ApprovedLoans.options.length; i++){
				document.frm0101.PriorityList.value = document.frm0101.PriorityList.value + ":" + document.frm0101.ApprovedLoans.options[i].value;
			}
		}
	}
	
	function ViewLoan(){
		loan_id = document.frm0101.ApprovedLoans.value;
		if (loan_id > 0) {
			openWindow('m008FS3.asp?intLoan_req_id='+loan_id,'');
		}
	}

	function Move(dir){
		if (document.frm0101.ApprovedLoans.selectedIndex == -1) {
			alert("Select a loan.");
			return ;
		}
		var temp_value = "";
		var temp_option = "";
		switch (dir) {
			case 'up':
				if (document.frm0101.ApprovedLoans.selectedIndex == 0) break;
				temp_text = document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex].text;
				temp_value = document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex].value;
				document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex].text = document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex-1].text
				document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex].value = document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex-1].value
				document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex-1].text = temp_text;
				document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex-1].value = temp_value;				
				document.frm0101.ApprovedLoans.selectedIndex = document.frm0101.ApprovedLoans.selectedIndex - 1;
			break;
			case 'down':
				if ((document.frm0101.ApprovedLoans.selectedIndex + 1) == document.frm0101.ApprovedLoans.options.length) break;			
				temp_text = document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex].text;
				temp_value = document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex].value;
				document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex].text = document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex+1].text
				document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex].value = document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex+1].value
				document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex+1].text = temp_text;
				document.frm0101.ApprovedLoans.options[document.frm0101.ApprovedLoans.selectedIndex+1].value = temp_value;	
				document.frm0101.ApprovedLoans.selectedIndex = document.frm0101.ApprovedLoans.selectedIndex + 1;				
			break;			
		}
	}
	</script>
</head>
<body onload="">
<form name="frm0101" method="POST" action="m008c0101.asp">
<h3>Work Priority</h3>
<i>Ordered from highest to lowest priority.  Double click a loan to view details.</i>  
<hr>
<select name="ApprovedLoans" size="20" style="width: 300px;" ondblclick="ViewLoan();" tabindex="1" accesskey="F">
<%
var count = 0;
while (!rsLoans.EOF) {
	if (rsLoans.Fields.Item("intLoan_req_id").Value > 0) {
%>
	<option value="<%=rsLoans.Fields.Item("intLoan_req_id").Value%>"><%=Trim(rsLoans.Fields.Item("chvLoan_name").Value)%>		
<%
		count++;
	}
	rsLoans.MoveNext();
}
%>
</select>
<% 
if ( Session("MM_UserAuthorization") >= 6 ) { 
%>
<table cellpadding="1" cellspacing="1" width="300">
	<tr>
		<td align="center">
			<input type="button" name="btnMoveUp" value="Move Up" onClick="Move('up');" class="btnstyle" tabindex="2">&nbsp;
			<input type="button" name="btnMoveDown" value="Move Down" onClick="Move('down');" class="btnstyle" tabindex="3">
		</td>
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" class="btnstyle" tabindex="4"></td>
	</tr>
</table>
<%
}
%>
<input type="hidden" name="Count" value="<%=count%>">
<input type="hidden" name="PriorityList" value="">
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsLoans.Close();
%>
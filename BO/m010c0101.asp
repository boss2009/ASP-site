<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.Form("MM_update")) == "true") {
	var BuyoutArray
	BuyoutArray = String(Request.Form("PriorityList")).split(":");
	if (BuyoutArray.length==1) {
		var rsRankBuyout = Server.CreateObject("ADODB.Recordset");
		rsRankBuyout.ActiveConnection = MM_cnnASP02_STRING;
		rsRankBuyout.Source = "{call dbo.cp_update_Buyout_req_rank("+BuyoutArray+",0,0)}";
		rsRankBuyout.CursorType = 0;
		rsRankBuyout.CursorLocation = 2;
		rsRankBuyout.LockType = 3;
		rsRankBuyout.Open();	
	}
	if (BuyoutArray.length > 1) {
		var rsRankBuyout = Server.CreateObject("ADODB.Recordset");
		rsRankBuyout.ActiveConnection = MM_cnnASP02_STRING;
		rsRankBuyout.Source = "{call dbo.cp_update_Buyout_req_rank("+BuyoutArray[0]+",0,0)}";
		rsRankBuyout.CursorType = 0;
		rsRankBuyout.CursorLocation = 2;
		rsRankBuyout.LockType = 3;
		rsRankBuyout.Open();	
		for (var i = 1; i < BuyoutArray.length; i ++) {
			rsRankBuyout.Source = "{call dbo.cp_update_Buyout_req_rank("+BuyoutArray[i]+","+i+",0)}";
			rsRankBuyout.Open();
		}		
	}
}

var rsBuyouts = Server.CreateObject("ADODB.Recordset");
rsBuyouts.ActiveConnection = MM_cnnASP02_STRING;
rsBuyouts.Source = "{call dbo.cp_Buyout_request3(0,0,0,'',0,'',0,0,0,0,2,'Q',0)}";
rsBuyouts.CursorType = 0;
rsBuyouts.CursorLocation = 2;
rsBuyouts.LockType = 3;
rsBuyouts.Open();
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
			document.frm0101.PriorityList.value = document.frm0101.ApprovedBuyouts.options[0].value;
			for (var i = 1; i < document.frm0101.ApprovedBuyouts.options.length; i++){
				document.frm0101.PriorityList.value = document.frm0101.PriorityList.value + ":" + document.frm0101.ApprovedBuyouts.options[i].value;
			}
		}
	}
	
	function ViewBuyout(){
		Buyout_id = document.frm0101.ApprovedBuyouts.value;
		if (Buyout_id > 0) {
			openWindow('m010FS3.asp?intBuyout_req_id='+Buyout_id,'');
		}
	}

	function Move(dir){
		if (document.frm0101.ApprovedBuyouts.selectedIndex == -1) {
			alert("Select a Buyout.");
			return ;
		}
		var temp_value = "";
		var temp_option = "";
		switch (dir) {
			case 'up':
				if (document.frm0101.ApprovedBuyouts.selectedIndex == 0) break;
				temp_text = document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex].text;
				temp_value = document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex].value;
				document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex].text = document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex-1].text
				document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex].value = document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex-1].value
				document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex-1].text = temp_text;
				document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex-1].value = temp_value;				
				document.frm0101.ApprovedBuyouts.selectedIndex = document.frm0101.ApprovedBuyouts.selectedIndex - 1;
			break;
			case 'down':
				if ((document.frm0101.ApprovedBuyouts.selectedIndex + 1) == document.frm0101.ApprovedBuyouts.options.length) break;			
				temp_text = document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex].text;
				temp_value = document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex].value;
				document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex].text = document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex+1].text
				document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex].value = document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex+1].value
				document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex+1].text = temp_text;
				document.frm0101.ApprovedBuyouts.options[document.frm0101.ApprovedBuyouts.selectedIndex+1].value = temp_value;	
				document.frm0101.ApprovedBuyouts.selectedIndex = document.frm0101.ApprovedBuyouts.selectedIndex + 1;				
			break;			
		}
	}
	</script>
</head>
<body onload="">
<form name="frm0101" method="POST" action="m010c0101.asp">
<h3>Work Priority</h3>
<i>Ordered from highest to lowest priority.  Double click a buyout to view details.</i>  
<hr>
<select name="ApprovedBuyouts" size="20" style="width: 300px;" ondblclick="ViewBuyout();" tabindex="1" accesskey="F">
<%
var count = 0;
while (!rsBuyouts.EOF) {
	if (rsBuyouts.Fields.Item("intBuyout_req_id").Value > 0) {
%>
	<option value="<%=rsBuyouts.Fields.Item("intBuyout_req_id").Value%>"><%=Trim(rsBuyouts.Fields.Item("intBuyout_req_id").Value)%>		
<%
		count++;
	}
	rsBuyouts.MoveNext();
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
rsBuyouts.Close();
%>
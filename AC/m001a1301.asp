<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_action")) == "insert") {
	var Equipment = String(Request.Form("Equipment")).replace(/'/g, "''");
	var rsTraining = Server.CreateObject("ADODB.Recordset");
	rsTraining.ActiveConnection = MM_cnnASP02_STRING;
	rsTraining.Source = "{call dbo.cp_ac_training_request(0,"+Request.QueryString("intAdult_id")+",'"+CurrentDate()+"','"+Equipment+"','"+Request.Form("TrainingStatus")+"','"+Request.Form("Date")+"',"+Request.Form("Reason")+","+Request.Form("TrainedBy")+","+Request.Form("Hours")+","+Request.Form("Minutes")+",0,'A',0)}";
	rsTraining.CursorType = 0;
	rsTraining.CursorLocation = 2;
	rsTraining.LockType = 3;
	rsTraining.Open();	
	Response.Redirect("InsertSuccessful.html");
}

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsReason = Server.CreateObject("ADODB.Recordset");
rsReason.ActiveConnection = MM_cnnASP02_STRING;
rsReason.Source = "{call dbo.cp_doc_cdn_rsn(0,'5','',2,'Q',0)}";
rsReason.CursorType = 0;
rsReason.CursorLocation = 2;
rsReason.LockType = 3;
rsReason.Open();
%>
<html>
<head>
	<title>New Training Request</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init(){
		ChangeTrainingStatus();	
		document.frm1301.TrainingStatus.focus();
	}

	function ChangeTrainingStatus(){
		Reason.style.visibility = "hidden";		
		TrainedBy.style.visibility = "hidden";			
		TrainedTime.style.visibility = "hidden";
		Hours.style.visibility = "hidden";		
		Minutes.style.visibility = "hidden";
		
		document.frm1301.Reason.style.visibility = "hidden";		
		document.frm1301.TrainedBy.style.visibility = "hidden";					
		document.frm1301.Hours.style.visibility = "hidden";		
		document.frm1301.Minutes.style.visibility = "hidden";			
		
		switch (document.frm1301.TrainingStatus.value) {
			case "0":
				Reason.style.visibility = "visible";
				document.frm1301.Reason.style.visibility = "visible";						
			break;
			case "2":
				TrainedTime.style.visibility = "visible";
				TrainedBy.style.visibility = "visible";				
				Hours.style.visibility = "visible";		
				Minutes.style.visibility = "visible";
				document.frm1301.TrainedBy.style.visibility = "visible";				
				document.frm1301.Hours.style.visibility = "visible";		
				document.frm1301.Minutes.style.visibility = "visible";						
			break;
		}
	}
		
	function Save(){
		if (!CheckDate(document.frm1301.Date.value)) {
			alert("Invalid Date.");
			document.frm1301.Date.focus();
			return ;
		}
		document.frm1301.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm1301">
<h5>New Training Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Equipment:</td>
		<td><input type="text" name="Equipment" accesskey="F" tabindex="1" size="65"></td>
	</tr>
	<tr>
		<td>Training Status:</td>
		<td><select name="TrainingStatus" tabindex="2" onChange="ChangeTrainingStatus();">
				<option value="0">Unable to arrange
				<option value="1">Declined
				<option value="2" SELECTED>Completed
		</select></td>
	</tr>
	<tr>
		<td>Date:</td>
		<td>
			<input type="text" name="Date" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="3" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td><span id="Reason">Reason:</span></td>
		<td><select name="Reason" tabindex="4">
			<%
			while (!rsReason.EOF) {			
			%>
				<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>"><%=rsReason.Fields.Item("chvDocDesc").Value%>
			<%
				rsReason.MoveNext();
			}
			%>			
		</select></td>
	</tr>	
	<tr> 
		<td><span id="TrainedBy">Trained By:</span></td>
		<td><select name="TrainedBy" tabindex="5">
			<% 
			var staffid = Session("insStaff_id");
			while (!rsStaff.EOF) {
			%>
				<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==staffid)?"SELECTED":"")%>><%=(rsStaff.Fields.Item("chvName").Value)%></option>
			<%
				rsStaff.MoveNext();
			}
			%>
        </select></td>		
    </tr>
	<tr>
		<td><span id="TrainedTime">Trained Time:</span></td>
		<td>
			<select name="Hours" tabindex="6">
				<option value="1">0			
				<option value="1" SELECTED>1
				<option value="2">2
				<option value="3">3
				<option value="4">4
				<option value="5">5
				<option value="6">6
				<option value="7">7
				<option value="8">8							
			</select>
			<span id="Hours">Hours</span>
			<select name="Minutes" tabindex="7">
				<option value="0">0
				<option value="15">15
				<option value="30">30
				<option value="45">45
			</select>
			<span id="Minutes">Minutes</span>		
		</td>
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="8" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="9" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="insert">
</form>
</body>
</html>
<%
rsStaff.Close();
rsReason.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_action")) == "update") {
	var rsTraining = Server.CreateObject("ADODB.Recordset");
	rsTraining.ActiveConnection = MM_cnnASP02_STRING;
	rsTraining.Source = "{call dbo.cp_buyout_training_status4("+Request.Form("MM_recordId")+",0,'"+Request.Form("TrainingStatus")+"','"+Request.Form("Date")+"',"+Request.Form("Reason")+","+Request.Form("TrainedBy")+","+Request.Form("Hours")+","+Request.Form("Minutes")+","+Session("insStaff_id")+",0,'E',0)}";
	rsTraining.CursorType = 0;
	rsTraining.CursorLocation = 2;
	rsTraining.LockType = 3;	
	rsTraining.Open();	
	
	//Trigger to insert a service request into client service and notes if the training status is Completed
	if (String(Request.Form("InsertService"))=="True") {
		var EquipList = "";
		var rsEquipment = Server.CreateObject("ADODB.Recordset");
		rsEquipment.ActiveConnection = MM_cnnASP02_STRING;
		rsEquipment.Source = "{call dbo.cp_buyout_eqp_requested(0,"+Request.QueryString("intBuyout_req_id")+",0,0,0,0.0,0,'',0,'Q',0)}";
		rsEquipment.CursorType = 0;
		rsEquipment.CursorLocation = 2;
		rsEquipment.LockType = 3;
		rsEquipment.Open();
		while (!rsEquipment.EOF) {
			if (rsEquipment.Fields.Item("bitIs_Train_request").Value=="1") {
				if (rsEquipment.Fields.Item("bitIs_class").Value=="1") {
					EquipList = EquipList + rsEquipment.Fields.Item("chv_Eqp_Class_Name").Value + "\n";			
				} else {
					EquipList = EquipList + rsEquipment.Fields.Item("chvBundle_Name").Value + "\n";							
				}
			}
			rsEquipment.MoveNext();
		}
		rsEquipment.Close();
		
		var rsUserType = Server.CreateObject("ADODB.Recordset");
		rsUserType.ActiveConnection = MM_cnnASP02_STRING;
		rsUserType.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
		rsUserType.CursorType = 0;
		rsUserType.CursorLocation = 2;
		rsUserType.LockType = 3;
		rsUserType.Open();
	
		var Year = CurrentYear();
		var Cycle = CurrentMonth();

		var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
		rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
		rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsUserType.Fields.Item("intEq_user_id").Value+",0,'"+Request.Form("Date")+"',"+Year+","+Cycle+","+Request.Form("TrainedBy")+",'"+String(EquipList).replace(/'/g, "''")+"','3400000000000000000000000000000000000000',0,'A',0)}";
		rsServiceRequested.CursorType = 0;
		rsServiceRequested.CursorLocation = 2;
		rsServiceRequested.LockType = 3;
		rsServiceRequested.Open();
	}		
	
	//Trigger to insert a note into client services and notes if the training status is Declined or Incomplete
	if (String(Request.Form("InsertNote"))=="True") {
		var notes = "re: ";
		if (String(Request.Form("TrainingStatus"))=="1") {
			notes = notes + "Declined\n";
		} else {
			notes = notes + "Unable to arrange, ";
			var rsReason = Server.CreateObject("ADODB.Recordset");
			rsReason.ActiveConnection = MM_cnnASP02_STRING;
			rsReason.Source = "{call dbo.cp_doc_cdn_rsn2("+Request.Form("Reason")+",0,'',1,'Q',0)}";
			rsReason.CursorType = 0;
			rsReason.CursorLocation = 2;
			rsReason.LockType = 3;			
			rsReason.Open();
			if (!rsReason.EOF) notes = notes + rsReason.Fields.Item("chvDocDesc").Value			
		}
		
		var rsUserType = Server.CreateObject("ADODB.Recordset");
		rsUserType.ActiveConnection = MM_cnnASP02_STRING;
		rsUserType.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
		rsUserType.CursorType = 0;
		rsUserType.CursorLocation = 2;
		rsUserType.LockType = 3;
		rsUserType.Open();

		var rsNote = Server.CreateObject("ADODB.Recordset");		
		rsNote.ActiveConnection = MM_cnnASP02_STRING;
		rsNote.Source = "{call dbo.cp_ac_srv_note("+rsUserType.Fields.Item("intEq_user_id").Value+",0,'"+Request.Form("Date")+"',0,0,"+Request.Form("TrainedBy")+",'"+String(notes).replace(/'/g, "''")+"','0500000000000000000000000000000000000000',1,'A',0)}";
		rsNote.CursorType = 0;
		rsNote.CursorLocation = 2;
		rsNote.LockType = 3;
		rsNote.Open();	
	}
	
	Response.Redirect("m010e0702.asp?intBuyout_Req_id="+Request.QueryString("intBuyout_Req_id"));
}

var rsTraining = Server.CreateObject("ADODB.Recordset");
rsTraining.ActiveConnection = MM_cnnASP02_STRING;
rsTraining.Source = "{call dbo.cp_Buyout_training_status4(0,"+Request.QueryString("intBuyout_Req_id")+",'','',0,0,0,0,0,0,'Q',0)}";
rsTraining.CursorType = 0;
rsTraining.CursorLocation = 2;
rsTraining.LockType = 3;
rsTraining.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsReason = Server.CreateObject("ADODB.Recordset");
rsReason.ActiveConnection = MM_cnnASP02_STRING;
rsReason.Source = "{call dbo.cp_doc_cdn_rsn2(0,'5','',2,'Q',0)}";
rsReason.CursorType = 0;
rsReason.CursorLocation = 2;
rsReason.LockType = 3;
rsReason.Open();
%>
<html>
<head>
	<title>Client Training Status</title>
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
				window.location.href='m010e0701.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>';
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init(){
	<%
	if (!rsTraining.EOF) {
	%>
		ChangeTrainingStatus();	
		document.frm0702.TrainingStatus.focus();
	<%
	}
	%>
	}

	function ChangeTrainingStatus(){
		Reason.style.visibility = "hidden";		
		TrainedBy.style.visibility = "hidden";			
		TrainedTime.style.visibility = "hidden";
		Hours.style.visibility = "hidden";		
		Minutes.style.visibility = "hidden";
		
		document.frm0702.Reason.style.visibility = "hidden";		
		document.frm0702.TrainedBy.style.visibility = "hidden";					
		document.frm0702.Hours.style.visibility = "hidden";		
		document.frm0702.Minutes.style.visibility = "hidden";			
		
		switch (document.frm0702.TrainingStatus.value) {
			case "0":
				Reason.style.visibility = "visible";
				document.frm0702.Reason.style.visibility = "visible";						
			break;
			case "2":
				TrainedTime.style.visibility = "visible";
				TrainedBy.style.visibility = "visible";				
				Hours.style.visibility = "visible";		
				Minutes.style.visibility = "visible";
				document.frm0702.TrainedBy.style.visibility = "visible";				
				document.frm0702.Hours.style.visibility = "visible";		
				document.frm0702.Minutes.style.visibility = "visible";						
			break;
		}
	}
		
	function Save(){
		if (!CheckDate(document.frm0702.Date.value)) {
			alert("Invalid Date.");
			document.frm0702.Date.focus();
			return ;
		}
		
		if ((document.frm0702.TrainingStatus.value=="2") && (document.frm0702.InsertService.value=="True")) {
			document.frm0702.InsertService.value="True"
		} else {
			document.frm0702.InsertService.value="False"		
		}

		if (((document.frm0702.TrainingStatus.value=="0") || (document.frm0702.TrainingStatus.value=="1")) && (document.frm0702.InsertNote.value=="True")) {
			document.frm0702.InsertNote.value="True"
		} else {
			document.frm0702.InsertNote.value="False"		
		}
		
		document.frm0702.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0702">
<h5>Client Training Status</h5>
<hr>
<%
if (rsTraining.EOF) {
%>
<i>Not available without training request.</i>
<%
} else {
%>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Training Status</td>
		<td><select name="TrainingStatus" accesskey="F" tabindex="1" onChange="ChangeTrainingStatus();">
				<option value="0" <%=((rsTraining.Fields.Item("chrClnt_Trn_Sts").Value=="0")?"SELECTED":"")%>>Unable to arrange
				<option value="1" <%=((rsTraining.Fields.Item("chrClnt_Trn_Sts").Value=="1")?"SELECTED":"")%>>Declined
				<option value="2" <%=((rsTraining.Fields.Item("chrClnt_Trn_Sts").Value=="2")?"SELECTED":"")%>>Completed
		</select></td>
	</tr>
	<tr>
		<td>Date:</td>
		<td>
			<input type="text" name="Date" value="<%=FilterDate(rsTraining.Fields.Item("dtsClnt_date").Value)%>" size="11" maxlength="10" tabindex="2" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td><span id="Reason">Reason:</span></td>
		<td><select name="Reason" tabindex="3">
			<%
			while (!rsReason.EOF) {			
			%>
				<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>" <%=((rsReason.Fields.Item("intDoc_id").Value==rsTraining.Fields.Item("insClnt_Rjt_Rsn_id").value)?"SELECTED":"")%>><%=rsReason.Fields.Item("chvDocDesc").Value%>
			<%
				rsReason.MoveNext();
			}
			%>			
		</select></td>
	</tr>	
	<tr> 
		<td><span id="TrainedBy">Trained By:</span></td>
		<td><select name="TrainedBy" tabindex="4">
			<% 
			var staffid = Session("insStaff_id");
			if (rsTraining.Fields.Item("insClnt_Trn_Staff_id").Value != null) staffid = rsTraining.Fields.Item("insClnt_Trn_Staff_id").Value;
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
			<select name="Hours" tabindex="5">
				<option value="1" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==0)?"SELECTED":"")%>>0			
				<option value="1" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==1)?"SELECTED":"")%>>1
				<option value="2" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==2)?"SELECTED":"")%>>2
				<option value="3" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==3)?"SELECTED":"")%>>3
				<option value="4" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==4)?"SELECTED":"")%>>4
				<option value="5" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==5)?"SELECTED":"")%>>5
				<option value="6" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==6)?"SELECTED":"")%>>6
				<option value="7" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==7)?"SELECTED":"")%>>7
				<option value="8" <%=((rsTraining.Fields.Item("insClnt_Trn_hr").Value==8)?"SELECTED":"")%>>8							
			</select>
			<span id="Hours">Hours</span>
			<select name="Minutes" tabindex="6">
				<option value="0" <%=((rsTraining.Fields.Item("insClnt_Trn_min").Value==0)?"SELECTED":"")%>>0
				<option value="15" <%=((rsTraining.Fields.Item("insClnt_Trn_min").Value==15)?"SELECTED":"")%>>15
				<option value="30" <%=((rsTraining.Fields.Item("insClnt_Trn_min").Value==30)?"SELECTED":"")%>>30
				<option value="45" <%=((rsTraining.Fields.Item("insClnt_Trn_min").Value==45)?"SELECTED":"")%>>45
			</select>
			<span id="Minutes">Minutes</span>		
		</td>
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="7" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="8" onClick="window.location.href='m010e0701.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="update">
<input type="hidden" name="InsertService" value="<%=((rsTraining.Fields.Item("chrClnt_Trn_Sts").Value==null)?"True":"False")%>">
<input type="hidden" name="InsertNote" value="<%=((rsTraining.Fields.Item("chrClnt_Trn_Sts").Value==null)?"True":"False")%>">
<input type="hidden" name="MM_recordId" value="<%=rsTraining.Fields.Item("intTrainReq_Id").Value%>">
<%
}
%>
</form>
</body>
</html>
<%
rsTraining.Close();
rsStaff.Close();
rsReason.Close();
%>
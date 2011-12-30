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
	rsTraining.Source = "{call dbo.cp_loan_training_request3("+Request.Form("MM_recordId")+",0,"+Request.Form("TrainingStatus")+","+Request.Form("ServiceProvider")+",0,'E',0)}";
	rsTraining.CursorType = 0;
	rsTraining.CursorLocation = 2;
	rsTraining.LockType = 3;
	rsTraining.Open();

	var rsTrainingRequired = Server.CreateObject("ADODB.Recordset");
	rsTrainingRequired.ActiveConnection = MM_cnnASP02_STRING;
	rsTrainingRequired.CursorType = 0;
	rsTrainingRequired.CursorLocation = 2;
	rsTrainingRequired.LockType = 3;
	
	var rsEquipment = Server.CreateObject("ADODB.Recordset");
	rsEquipment.ActiveConnection = MM_cnnASP02_STRING;
	rsEquipment.Source = "{call dbo.cp_eqp_requested(0,"+Request.QueryString("intLoan_Req_id")+",0,0,0,'',0.0,0,0,'Q',0)}";
	rsEquipment.CursorType = 0;
	rsEquipment.CursorLocation = 2;
	rsEquipment.LockType = 3;
	rsEquipment.Open();
	while (!rsEquipment.EOF) {
		rsTrainingRequired.Source = "{call dbo.cp_update_eqp_requested(" + rsEquipment.Fields.Item("intEqpReq_Id").Value + ",0,1)}";	
		rsTrainingRequired.Open();
		rsEquipment.MoveNext();
	}
	
	if (Request.Form("TrainingRequired").Count > 0) {
		if (Request.Form("TrainingRequired").Count > 1) {
			for (var i = 1; i <= Request.Form("TrainingRequired").Count; i++) {
				rsTrainingRequired.Source = "{call dbo.cp_update_eqp_requested(" + Request.Form("TrainingRequired")(i) + ",1,1)}";	
				rsTrainingRequired.Open();
			}
		} else {
			rsTrainingRequired.Source = "{call dbo.cp_update_eqp_requested(" + Request.Form("TrainingRequired") + ",1,1)}";	
			rsTrainingRequired.Open();
		}
	}
	Response.Redirect("m008e0601.asp?intLoan_Req_id="+Request.QueryString("intLoan_Req_id"));
}

if (String(Request("MM_action")) == "insert") {
	var rsTraining = Server.CreateObject("ADODB.Recordset");
	rsTraining.ActiveConnection = MM_cnnASP02_STRING;
	rsTraining.Source = "{call dbo.cp_loan_training_request3(0,"+Request.QueryString("intLoan_req_id")+","+Request.Form("TrainingStatus")+","+Request.Form("ServiceProvider")+",0,'A',0)}";
	rsTraining.CursorType = 0;
	rsTraining.CursorLocation = 2;
	rsTraining.LockType = 3;
	rsTraining.Open();

	var rsTrainingRequired = Server.CreateObject("ADODB.Recordset");
	rsTrainingRequired.ActiveConnection = MM_cnnASP02_STRING;
	rsTrainingRequired.CursorType = 0;
	rsTrainingRequired.CursorLocation = 2;
	rsTrainingRequired.LockType = 3;

	if (Request.Form("TrainingRequired").Count > 0) {
		if (Request.Form("TrainingRequired").Count > 1) {
			for (var i = 1; i <= Request.Form("TrainingRequired").Count; i++) {
				rsTrainingRequired.Source = "{call dbo.cp_update_eqp_requested(" + Request.Form("TrainingRequired")(i) + ",1,1)}";	
				rsTrainingRequired.Open();
			}
		} else {
			rsTrainingRequired.Source = "{call dbo.cp_update_eqp_requested(" + Request.Form("TrainingRequired") + ",1,1)}";	
			rsTrainingRequired.Open();
		}
	}
	Response.Redirect("m008e0601.asp?intLoan_Req_id="+Request.QueryString("intLoan_Req_id"));	
}

var rsTraining = Server.CreateObject("ADODB.Recordset");
rsTraining.ActiveConnection = MM_cnnASP02_STRING;
rsTraining.Source = "{call dbo.cp_loan_training_request3(0,"+Request.QueryString("intLoan_req_id")+",0,0,0,'Q',0)}";
rsTraining.CursorType = 0;
rsTraining.CursorLocation = 2;
rsTraining.LockType = 3;
rsTraining.Open();

var IsNew = ((rsTraining.EOF)?true:false);

var rsServiceProvider = Server.CreateObject("ADODB.Recordset");
rsServiceProvider.ActiveConnection = MM_cnnASP02_STRING;
rsServiceProvider.Source = "{call dbo.cp_srv_pvdr(0,'',0,'Q',0)}";
rsServiceProvider.CursorType = 0;
rsServiceProvider.CursorLocation = 2;
rsServiceProvider.LockType = 3;
rsServiceProvider.Open();

var rsEquipment = Server.CreateObject("ADODB.Recordset");
rsEquipment.ActiveConnection = MM_cnnASP02_STRING;
rsEquipment.Source = "{call dbo.cp_eqp_requested(0,"+Request.QueryString("intLoan_Req_id")+",0,0,0,'',0.0,0,0,'Q',0)}";
rsEquipment.CursorType = 0;
rsEquipment.CursorLocation = 2;
rsEquipment.LockType = 3;
rsEquipment.Open();
%>
<html>
<head>
	<title>Training Requested</title>
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
			case 85 :
				//alert("U");
				document.frm0601.reset();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init() {
		document.frm0601.TrainingStatus.focus();
	}

	function openWindow(page) {
		if (page!='nothing') win1=window.open(page, "", "width=300,height=300,scrollbars=1,left=300,top=300,status=1");
		return ;
	}
	
	function Save(){
		document.frm0601.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0601">
<h5>Training Requested</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Training Status:</td>
		<td nowrap><select name="TrainingStatus" tabindex="1" accesskey="F">
			<option value="1" <%if (!IsNew) Response.Write(((rsTraining.Fields.Item("bitIs_active").Value=="1")?"SELECTED":""))%>>Training Requested
			<option value="0" <%if (!IsNew) Response.Write(((rsTraining.Fields.Item("bitIs_active").Value=="0")?"SELECTED":""))%>>Not Available
		</select></td>		
	</tr>  
	<tr> 
		<td nowrap>Service Provider:</td>
		<td nowrap><select name="ServiceProvider" tabindex="2">
			<% 
			while (!rsServiceProvider.EOF) {
			%>
				<option value="<%=(rsServiceProvider.Fields.Item("intSPvdr_id").Value)%>"  <%if (!IsNew) Response.Write(((rsTraining.Fields.Item("insSrv_pvdr_id").Value==rsServiceProvider.Fields.Item("intSPvdr_id").Value)?"SELECTED":""))%>><%=(rsServiceProvider.Fields.Item("chvSPvdr_Desc").Value)%></option>
			<%
				rsServiceProvider.MoveNext();
			}
			%>
		</select></td>		
	</tr>
</table>
<div class="BrowsePanel" style="width: 445px; height: 180px; top: 115px;"> 
<table cellspacing="1" cellpadding="1">
	<tr> 
        <th class="headrow" width="300" align="left" nowrap>Equipment Name</th>
        <th class="headrow" width="100" align="center" nowrap>Training Required</th>
	</tr>
<%
while (!rsEquipment.EOF) {
%>
	<tr> 
		<td valign="top" align="left"><input type="text" name="EquipmentName" style="border: none;background-color: #ffffe6;" value="<%=((rsEquipment.Fields.Item("bitIs_class").Value=="1")?rsEquipment.Fields.Item("chvEqp_Class_Name").Value:rsEquipment.Fields.Item("chvEqp_Bundle_Name").Value)%>" readonly size="48"></td>
        <td valign="top" align="center"><input type="checkbox" name="TrainingRequired" style="background-color: #ffffe6;" value="<%=rsEquipment.Fields.Item("intEqpReq_Id").Value%>" <%=((rsEquipment.Fields.Item("bitIs_Train_request").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
	</tr>
<%
	rsEquipment.MoveNext();
}
%>
</table>
</div>
<div style="position: absolute; top: 300px">
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" class="btnstyle"></td>
    </tr>
</table>
</div>
<input type="hidden" name="MM_action" value="<%=((IsNew)?"insert":"update")%>">
<input type="hidden" name="MM_recordId" value="<%=((IsNew)?0:rsTraining.Fields.Item("intTrainReq_Id").value)%>">
</form>
</body>
</html>
<%
rsTraining.Close();
rsServiceProvider.Close();
rsEquipment.Close();
%>
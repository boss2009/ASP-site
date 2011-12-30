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
	rsTraining.Source = "{call dbo.cp_buyout_training_request4("+Request.Form("MM_recordId")+","+Request.QueryString("intBuyout_req_id")+","+Request.Form("TrainingStatus")+","+Request.Form("ServiceProvider")+","+Session("insStaff_id")+",0,'E',0)}";
	rsTraining.CursorType = 0;
	rsTraining.CursorLocation = 2;
	rsTraining.LockType = 3;
	rsTraining.Open();

	var rsTrainingRequired = Server.CreateObject("ADODB.Recordset");
	rsTrainingRequired.ActiveConnection = MM_cnnASP02_STRING;
	rsTrainingRequired.CursorType = 0;
	rsTrainingRequired.CursorLocation = 2;
	rsTrainingRequired.LockType = 3;
	
	var rsEquipmentRequested = Server.CreateObject("ADODB.Recordset");
	rsEquipmentRequested.ActiveConnection = MM_cnnASP02_STRING;
	rsEquipmentRequested.Source = "{call dbo.cp_buyout_eqp_requested(0,"+Request.QueryString("intBuyout_req_id")+",0,0,0,0.0,0,'',0,'Q',0)}";
	rsEquipmentRequested.CursorType = 0;
	rsEquipmentRequested.CursorLocation = 2;
	rsEquipmentRequested.LockType = 3;
	rsEquipmentRequested.Open();
	while (!rsEquipmentRequested.EOF) {
		rsTrainingRequired.Source = "{call dbo.cp_update_BO_eqp_rqst(" + rsEquipmentRequested.Fields.Item("insBO_Eqp_Rqst_id").Value + ",0,1)}";	
		rsTrainingRequired.Open();
		rsEquipmentRequested.MoveNext();
	}
	
	if (Request.Form("TrainingRequired").Count > 0) {
		if (Request.Form("TrainingRequired").Count > 1) {
			for (var i = 1; i <= Request.Form("TrainingRequired").Count; i++) {
				rsTrainingRequired.Source = "{call dbo.cp_update_BO_eqp_rqst(" + Request.Form("TrainingRequired")(i) + ",1,1)}";	
				rsTrainingRequired.Open();
			}
		} else {
			rsTrainingRequired.Source = "{call dbo.cp_update_BO_eqp_rqst(" + Request.Form("TrainingRequired") + ",1,1)}";	
			rsTrainingRequired.Open();
		}
	}
	Response.Redirect("m010e0701.asp?intBuyout_Req_id="+Request.QueryString("intBuyout_Req_id"));
}

if (String(Request("MM_action")) == "insert") {
	var rsTraining = Server.CreateObject("ADODB.Recordset");
	rsTraining.ActiveConnection = MM_cnnASP02_STRING;
	rsTraining.Source = "{call dbo.cp_buyout_training_request4(0,"+Request.QueryString("intBuyout_Req_id")+","+Request.Form("TrainingStatus")+","+Request.Form("ServiceProvider")+","+Session("insStaff_id")+",0,'A',0)}";
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
				rsTrainingRequired.Source = "{call dbo.cp_update_BO_eqp_rqst(" + Request.Form("TrainingRequired")(i) + ",1,1)}";	
				rsTrainingRequired.Open();
			}
		} else {
			rsTrainingRequired.Source = "{call dbo.cp_update_BO_eqp_rqst(" + Request.Form("TrainingRequired") + ",1,1)}";	
			rsTrainingRequired.Open();
		}
	}
	Response.Redirect("m010e0701.asp?intBuyout_Req_id="+Request.QueryString("intBuyout_Req_id"));	
}

var rsTraining = Server.CreateObject("ADODB.Recordset");
rsTraining.ActiveConnection = MM_cnnASP02_STRING;
rsTraining.Source = "{call dbo.cp_buyout_training_request4(0,"+Request.QueryString("intBuyout_Req_id")+",0,0,0,0,'Q',0)}";
rsTraining.CursorType = 0;
rsTraining.CursorLocation = 2;
rsTraining.LockType = 3;
rsTraining.Open();

var IsNew = ((rsTraining.Fields.Item("bitIs_active").Value==null)?true:false);

var rsServiceProvider = Server.CreateObject("ADODB.Recordset");
rsServiceProvider.ActiveConnection = MM_cnnASP02_STRING;
rsServiceProvider.Source = "{call dbo.cp_srv_pvdr(0,'',0,'Q',0)}";
rsServiceProvider.CursorType = 0;
rsServiceProvider.CursorLocation = 2;
rsServiceProvider.LockType = 3;
rsServiceProvider.Open();

var rsEquipmentRequested = Server.CreateObject("ADODB.Recordset");
rsEquipmentRequested.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentRequested.Source = "{call dbo.cp_buyout_eqp_requested(0,"+Request.QueryString("intBuyout_req_id")+",0,0,0,0.0,0,'',0,'Q',0)}";
rsEquipmentRequested.CursorType = 0;
rsEquipmentRequested.CursorLocation = 2;
rsEquipmentRequested.LockType = 3;
rsEquipmentRequested.Open();
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
				document.frm0701.reset();
			break;
//		   	case 76 :
//				alert("L");
//				Close();
//			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init() {
		document.frm0701.TrainingStatus.focus();
	}

	function openWindow(page) {
		if (page!='nothing') win1=window.open(page, "", "width=300,height=300,scrollbars=1,left=300,top=300,status=1");
		return ;
	}
	
	function Save(){
		document.frm0701.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0701">
<h5>Training Requested</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Training Status:</td>
		<td><select name="TrainingStatus" tabindex="1" accesskey="F">
				<option value="1" <%if (!IsNew) Response.Write(((rsTraining.Fields.Item("bitIs_active").Value=="1")?"SELECTED":""))%>>Training Requested
				<option value="0" <%if (!IsNew) Response.Write(((rsTraining.Fields.Item("bitIs_active").Value=="0")?"SELECTED":""))%>>Not Available
		</select></td>		
	</tr>  
	<tr> 
		<td>Service Provider:</td>
		<td><select name="ServiceProvider" tabindex="2">
			<% 
			while (!rsServiceProvider.EOF) {
			%>
				<option value="<%=(rsServiceProvider.Fields.Item("intSPvdr_id").Value)%>"  <%if (!IsNew) Response.Write(((rsTraining.Fields.Item("insSrv_pvdr_id").Value==rsServiceProvider.Fields.Item("intSPvdr_id").Value)?"SELECTED":""))%>><%=(rsServiceProvider.Fields.Item("chvSPvdr_Desc").Value)%></option>
			<%
				rsServiceProvider.MoveNext();
			}
			%>
			</select>
		</td>		
	</tr>
</table>
<div class="BrowsePanel" style="height: 180px; top: 115px;"> 
<table cellspacing="1" cellpadding="0" border="0">
	<tr> 
		<th class="headrow" width="300" align="left">Equipment Name</th>
		<th class="headrow" nowrap align="center">Training Required</th>
	</tr>
<%
while (!rsEquipmentRequested.EOF) {
%>
	<tr> 
		<td align="left"><input type="text" name="EquipmentName" value="<%=((rsEquipmentRequested.Fields.Item("bitIs_class").Value=="1")?rsEquipmentRequested.Fields.Item("chv_Eqp_Class_Name").Value:rsEquipmentRequested.Fields.Item("chvBundle_Name").Value)%>" style="border: none;background-color: #ffffe6;" readonly size="48"></td>
		<td align="center"><input type="checkbox" name="TrainingRequired" value="<%=rsEquipmentRequested.Fields.Item("insBO_Eqp_Rqst_id").Value%>" <%=((rsEquipmentRequested.Fields.Item("bitIs_Train_request").Value=="1")?"CHECKED":"")%> class="chkstyle" style="background-color: #ffffe6;"></td>
	</tr>
<%
	rsEquipmentRequested.MoveNext();
}
%>
</table>
</div>
<div style="position: absolute; top: 300px">
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="" class="btnstyle"></td>
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
rsEquipmentRequested.Close();
%>
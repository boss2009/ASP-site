<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_action")) == "update"){
	var SecurityPlan = ((Request.Form("SecurityPlan")=="1")?"1":"0");
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	var EquipmentRequested = String(Request.Form("EquipmentRequested")).replace(/'/g, "''");				
	var rsReferringAgent = Server.CreateObject("ADODB.Recordset");
	rsReferringAgent.ActiveConnection = MM_cnnASP02_STRING;
	rsReferringAgent.Source = "{call dbo.cp_school_ref_subhdr2(" + Request.Form("ReferralID")+","+SecurityPlan+",'"+Request.Form("LoanFormReceivedDate")+"',"+Request.Form("ReferringAgent")+","+Request.Form("PILATStatus")+","+Request.Form("CaseManager")+",'"+EquipmentRequested+"','"+Notes+"','E',0)}";
	rsReferringAgent.CursorType = 0;
	rsReferringAgent.CursorLocation = 2;
	rsReferringAgent.LockType = 3;
	rsReferringAgent.Open();
	
	var rsReferralHardwareLocation = Server.CreateObject("ADODB.Recordset");
	rsReferralHardwareLocation.ActiveConnection = MM_cnnASP02_STRING;
	rsReferralHardwareLocation.CursorType = 0;
	rsReferralHardwareLocation.CursorLocation = 2;
	rsReferralHardwareLocation.LockType = 3;	
	//delete
	rsReferralHardwareLocation.Source = "{call dbo.cp_school_refhdr_loc(" + Request.QueryString("insSchool_id") + ",0,"+Request.Form("ReferralID")+",0,'D',0)}";
	rsReferralHardwareLocation.Open();
	//add
	for (var i=1; i<=Request.Form("Count"); i++){
		rsReferralHardwareLocation.Source = "{call dbo.cp_school_refhdr_loc(" + Request.QueryString("insSchool_id") + "," + Request.Form("Location")(i) + ","+Request.Form("ReferralID")+",0,'A',0)}";
		rsReferralHardwareLocation.Open();
	}	
	Response.Redirect("UpdateSuccessful.asp?page=m012e0202.asp&insSchool_id="+Request.QueryString("insSchool_id")+"&intReferral_id="+Request.QueryString("intReferral_id"));
}

if (String(Request.Form("MM_action")) == "insert") {
	var SecurityPlan = ((Request.Form("SecurityPlan")=="1")?"1":"0");	
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	var EquipmentRequested = String(Request.Form("EquipmentRequested")).replace(/'/g, "''");					
	var rsReferringAgent = Server.CreateObject("ADODB.Recordset");
	rsReferringAgent.ActiveConnection = MM_cnnASP02_STRING;
	rsReferringAgent.Source = "{call dbo.cp_school_ref_subhdr2(" + Request.Form("ReferralID")+","+SecurityPlan+",'"+Request.Form("LoanFormReceivedDate")+"',"+Request.Form("ReferringAgent")+","+Request.Form("PILATStatus")+","+Request.Form("CaseManager")+",'"+EquipmentRequested+"','"+Notes+"','A',0)}";
	rsReferringAgent.CursorType = 0;
	rsReferringAgent.CursorLocation = 2;
	rsReferringAgent.LockType = 3;
	rsReferringAgent.Open();
	
	var rsReferralHardwareLocation = Server.CreateObject("ADODB.Recordset");
	rsReferralHardwareLocation.ActiveConnection = MM_cnnASP02_STRING;
	rsReferralHardwareLocation.CursorType = 0;
	rsReferralHardwareLocation.CursorLocation = 2;
	rsReferralHardwareLocation.LockType = 3;	
	//delete
	rsReferralHardwareLocation.Source = "{call dbo.cp_school_refhdr_loc(" + Request.QueryString("insSchool_id") + ",0,"+Request.Form("ReferralID")+",0,'D',0)}";
	rsReferralHardwareLocation.Open();
	//add
	for (var i=1; i<=Request.Form("Count"); i++){
		rsReferralHardwareLocation.Source = "{call dbo.cp_school_refhdr_loc(" + Request.QueryString("insSchool_id") + "," + Request.Form("Location")(i) + ","+Request.Form("ReferralID")+",0,'A',0)}";
		rsReferralHardwareLocation.Open();
	}		
	Response.Redirect("m012e0202.asp?insSchool_id="+Request.QueryString("insSchool_id")+"&intReferral_id="+Request.QueryString("intReferral_id"));
}

var rsReferringAgent = Server.CreateObject("ADODB.Recordset");
rsReferringAgent.ActiveConnection = MM_cnnASP02_STRING;
rsReferringAgent.Source = "{call dbo.cp_school_ref_subhdr2(" + Request.QueryString("intReferral_id")+",0,'',0,0,0,'','','Q',0)}";
rsReferringAgent.CursorType = 0;
rsReferringAgent.CursorLocation = 2;
rsReferringAgent.LockType = 3;
rsReferringAgent.Open();

var IsNew = false;
if (rsReferringAgent.EOF) IsNew = true;

var rsReferralHardwareLocation = Server.CreateObject("ADODB.Recordset");
rsReferralHardwareLocation.ActiveConnection = MM_cnnASP02_STRING;
rsReferralHardwareLocation.Source = "{call dbo.cp_school_refhdr_loc("+Request.QueryString("insSchool_id")+",0,"+Request.QueryString("intReferral_id")+",2,'Q',0)}";
rsReferralHardwareLocation.CursorType = 0;
rsReferralHardwareLocation.CursorLocation = 2;
rsReferralHardwareLocation.LockType = 3;
rsReferralHardwareLocation.Open();

var count = 0;
while (!rsReferralHardwareLocation.EOF) {
	count++;
	rsReferralHardwareLocation.MoveNext();
}
if (count > 0) rsReferralHardwareLocation.MoveFirst();

var rsHardwareLocation = Server.CreateObject("ADODB.Recordset");
rsHardwareLocation.ActiveConnection = MM_cnnASP02_STRING;
rsHardwareLocation.Source = "{call dbo.cp_HW_Location(0,'',0,0,0,'Q',0)}";
rsHardwareLocation.CursorType = 0;
rsHardwareLocation.CursorLocation = 2;
rsHardwareLocation.LockType = 3;
rsHardwareLocation.Open();

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_school_contacts("+Request.QueryString("insSchool_id")+",0,0,0,'Q',0)}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_PILAT_status(0,'',0,'Q',0)}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();

var rsCaseManager = Server.CreateObject("ADODB.Recordset");
rsCaseManager.ActiveConnection = MM_cnnASP02_STRING;
rsCaseManager.Source = "{call dbo.cp_CaseMgr}";
rsCaseManager.CursorType = 0;
rsCaseManager.CursorLocation = 2;
rsCaseManager.LockType = 3;
rsCaseManager.Open();
%>
<html>
<head>
	<title>Update Referral Details</title>
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
			case 85:
				//alert("U");
				document.frm0202.reset();
			break;			
		   	case 76 :
				//alert("L");
				top.BodyFrame.location.href='m012q0201.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>';
			break;
		}
	}
	</script>
	<script language="JavaScript">
	var count = <%=count%>;	
		
	function Init() {
		ChangeSecurity();
		if (count > 3) count = 3;
		if (count > 0) {
			for (var i = 0; i < count; i++) {
				document.frm0202.Location[i].style.visibility = "visible";
			}
		}		
		if (count == 0) document.frm0202.Remove.disabled = true;
		if (count == 3) document.frm0202.Add.disabled = true;		
		document.frm0202.CaseManager.focus();
	}
		
	function Save(){
		if (!CheckTextArea(document.frm0202.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		if (!CheckDate(document.frm0202.LoanFormReceivedDate.value)){
			alert("Invalid Loan Form Received Date.");
			document.frm0202.LoanFormReceivedDate.focus();
			return ;
		}
		document.frm0202.Count.value = count;
		document.frm0202.submit();
	}
	
	function AddLocation(){
		document.frm0202.Location[count].style.visibility = "visible";
		count++;
		document.frm0202.Remove.disabled = false;
		if (count == 3) document.frm0202.Add.disabled = true;
	}
	
	function RemoveLocation(){
		count--;
		document.frm0202.Location[count].style.visibility = "hidden";
		document.frm0202.Add.disabled = false;
		if (count == 0) document.frm0202.Remove.disabled = true;
	}
	
	function ChangeSecurity(){
		if (document.frm0202.SecurityPlan.checked==true) {
			NotesLabel.style.visibility="visible";
			document.frm0202.Notes.style.visibility="visible";
		} else {
			NotesLabel.style.visibility="hidden";		
			document.frm0202.Notes.style.visibility="hidden";		
		}
	}
	</script>
</head>
<body onLoad="Init();">
<form action="<%=MM_updateAction%>" method="POST" name="frm0202">
<h5>Referral Details</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Case Manager:</td>
		<td nowrap><select name="CaseManager" tabindex="1" accesskey="F">
		<% 
		while (!rsCaseManager.EOF) {
		%>
			<option value="<%=(rsCaseManager.Fields.Item("insId").Value)%>"  <%if (IsNew) Response.Write(((rsCaseManager.Fields.Item("insId").Value==Session("insStaff_id"))?"SELECTED":""))%> <%if (!IsNew) Response.Write(((rsReferringAgent.Fields.Item("insCase_mngr_id").Value==rsCaseManager.Fields.Item("insId").Value)?"SELECTED":""))%>><%=(rsCaseManager.Fields.Item("chvName").Value)%>
		<%
			rsCaseManager.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>PILAT Status:</td>
		<td nowrap><select name="PILATStatus" tabindex="2">
		<% 
		while (!rsStatus.EOF) { 			
		%>
			<option value="<%=(rsStatus.Fields.Item("insPILAT_Status_id").Value)%>" <%if (IsNew) Response.Write(((rsStatus.Fields.Item("insPILAT_Status_id").Value=="1")?"SELECTED":""))%> <%if (!IsNew) Response.Write(((rsReferringAgent.Fields.Item("insPILAT_Status_id").Value==rsStatus.Fields.Item("insPILAT_Status_id").Value)?"SELECTED":""))%>><%=(rsStatus.Fields.Item("chvStatus_Desc").Value)%> 
		<% 
			rsStatus.MoveNext();
		} 
		%>		
		</select></td>
	</tr>
	<tr>
		<td nowrap>Referring Agent:</td>
		<td nowrap><select name="ReferringAgent" tabindex="3">
		<% 
		while (!rsContact.EOF) { 			
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%if (!IsNew) Response.Write(((rsReferringAgent.Fields.Item("intContact_id").Value==rsContact.Fields.Item("intContact_id").Value)?"SELECTED":""))%>><%=(rsContact.Fields.Item("chvFst_Name").Value)%>&nbsp;<%=(rsContact.Fields.Item("chvLst_Name").Value)%>&nbsp;<%=(rsContact.Fields.Item("chvRelationship").Value)%>
		<% 
			rsContact.MoveNext();
		} 
		%>		
			<option value="0">(none)			
		</select></td>	
	</tr>
	<tr>
		<td nowrap valign="top">
			Service/Equipment<br>
			Requested:<br>
			<i>(equipment/training/<br>
			consultation)</i>
		</td>
		<td nowrap valign="top"><textarea name="EquipmentRequested" Rows="5" cols="60" tabindex="4"><%=((!IsNew)?rsReferringAgent.Fields.Item("chvEqpList").Value:"")%></textarea></td>
	</tr>
	<tr>
		<td nowrap>Loan Form Received Date:</td>
		<td nowrap>
			<input type="text" name="LoanFormReceivedDate" value="<%=((!IsNew)?FilterDate(rsReferringAgent.Fields.Item("dtsSchool_Loan_form").Value):"")%>" size="11" maxlength="10" tabindex="5" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap valign="top">Hardware Location:</td>
		<td>
			<table cellpadding="1" cellspacing="1" style="border: 1px solid">
				<tr>
					<td><select name="Location" tabindex="6" style="visibility='hidden'">
					<%
					while (!rsHardwareLocation.EOF) {
						if (rsHardwareLocation.Fields.Item("bitIs_HW").Value=="1") {						
					%>
							<option value="<%=rsHardwareLocation.Fields.Item("insLocation_id").Value%>" <%if (!rsReferralHardwareLocation.EOF) { Response.Write(((rsHardwareLocation.Fields.Item("insLocation_id").Value==rsReferralHardwareLocation.Fields.Item("insLocation_id").Value)?"SELECTED":""));}%>><%=rsHardwareLocation.Fields.Item("chvLocation_Desc").Value%>
					<%
						}
						rsHardwareLocation.MoveNext();
					}
					rsHardwareLocation.MoveFirst();
					if (!rsReferralHardwareLocation.EOF) rsReferralHardwareLocation.MoveNext();
					%>
					</select></td>
				</tr>
				<tr>
					<td><select name="Location" tabindex="7" style="visibility='hidden'">
					<%
					while (!rsHardwareLocation.EOF) {
						if (rsHardwareLocation.Fields.Item("bitIs_HW").Value=="1") {						
					%>
							<option value="<%=rsHardwareLocation.Fields.Item("insLocation_id").Value%>" <%if (!rsReferralHardwareLocation.EOF) { Response.Write(((rsHardwareLocation.Fields.Item("insLocation_id").Value==rsReferralHardwareLocation.Fields.Item("insLocation_id").Value)?"SELECTED":""));}%>><%=rsHardwareLocation.Fields.Item("chvLocation_Desc").Value%>
					<%
						}
						rsHardwareLocation.MoveNext();
					}
					rsHardwareLocation.MoveFirst();
					if (!rsReferralHardwareLocation.EOF) rsReferralHardwareLocation.MoveNext();					
					%>
					</select></td>
				</tr>
				<tr>
					<td><select name="Location" tabindex="8" style="visibility='hidden'">
					<%
					while (!rsHardwareLocation.EOF) {
						if (rsHardwareLocation.Fields.Item("bitIs_HW").Value=="1") {						
					%>
							<option value="<%=rsHardwareLocation.Fields.Item("insLocation_id").Value%>" <%if (!rsReferralHardwareLocation.EOF) { Response.Write(((rsHardwareLocation.Fields.Item("insLocation_id").Value==rsReferralHardwareLocation.Fields.Item("insLocation_id").Value)?"SELECTED":""));}%>><%=rsHardwareLocation.Fields.Item("chvLocation_Desc").Value%>
					<%
						}
						rsHardwareLocation.MoveNext();
					}
					rsHardwareLocation.MoveFirst();
					if (!rsReferralHardwareLocation.EOF) rsReferralHardwareLocation.MoveNext();					
					%>
					</select></td>
				</tr>
			</table>				
		</td>
	</tr>
	<tr>
		<td></td>
		<td nowrap>
			<input type="button" name="Add" value="Add Location" onClick="AddLocation();" tabindex="9" class="btnstyle">&nbsp;
			<input type="button" name="Remove" value="Remove Location" onClick="RemoveLocation();" tabindex="10" class="btnstyle">					
		</td>
	</tr>	
    <tr> 
		<td nowrap>Security Plan:</td>
		<td nowrap><input type="checkbox" name="SecurityPlan" <%if (!IsNew) Response.Write(((rsReferringAgent.Fields.Item("bitIs_Security_Plan").Value=="1")?"CHECKED":""))%> value="1" tabindex="11" onClick="ChangeSecurity();" class="chkstyle"></td>
    </tr>	
	<tr>
		<td nowrap valign="top"><span id="NotesLabel">Notes on Security:</span></td>
		<td nowrap valign="top"><textarea name="Notes" Rows="3" cols="60" tabindex="12"><%=((!IsNew)?rsReferringAgent.Fields.Item("chvNote").Value:"")%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="13" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="14" class="btnstyle"></td>		
		<td><input type="button" value="Close" onClick="top.BodyFrame.location.href='m012q0201.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>'" tabindex="15" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ReferralID" value="<%=Request.QueryString("intReferral_id")%>">
<input type="hidden" name="MM_action" value="<%=((IsNew)?"insert":"update")%>">
<input type="hidden" name="Count" value="0">
</form>
</body>
</html>
<%
rsReferringAgent.Close();
rsHardwareLocation.Close();
rsContact.Close();
rsStatus.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_action")) == "update") {
	var Diagnosis = String(Request.Form("Diagnosis")).replace(/'/g, "''");
	var Specialist = String(Request.Form("Specialist")).replace(/'/g, "''");
	var Credentials = String(Request.Form("Credentials")).replace(/'/g, "''");
	var MedicalComments = String(Request.Form("MedicalComments")).replace(/'/g, "''");
	var CaseManagerComments = String(Request.Form("CaseManagerComments")).replace(/'/g, "''");
	var rsDisabilityDocumentation = Server.CreateObject("ADODB.Recordset");
	rsDisabilityDocumentation.ActiveConnection = MM_cnnASP02_STRING;
	rsDisabilityDocumentation.Source = "{call dbo.cp_disability_doc("+Request.Form("MM_recordId") + ","+ Request.QueryString("intAdult_id") + ","+Request.Form("TypeMedical")+","+Request.Form("TypeAudiology")+","+Request.Form("TypePsychoEd")+",'"+Request.Form("LocationOfDocumentation")+"','"+Request.Form("DocumentationDate")+"','"+Request.Form("DateReceived")+"','"+Diagnosis+"',"+Request.Form("Permanent")+",'"+Specialist+"','"+Credentials+"','"+MedicalComments+"',"+Request.Form("EligibleForASP")+",'"+Request.Form("PhoneAreaCode")+"','"+Request.Form("PhoneNumber")+"','"+Request.Form("PhoneExtension")+"','"+CaseManagerComments+"',0,'E',0)}";
	rsDisabilityDocumentation.CursorType = 0;
	rsDisabilityDocumentation.CursorLocation = 2;
	rsDisabilityDocumentation.LockType = 3;
	rsDisabilityDocumentation.Open();
  	Response.Redirect("UpdateSuccessful.asp?page=m001q0601.asp&intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsDocumentation = Server.CreateObject("ADODB.Recordset");
rsDocumentation.ActiveConnection = MM_cnnASP02_STRING;
rsDocumentation.Source = "{call dbo.cp_disability_doc("+ Request.QueryString("intDbtyDoc_id") + ",0,0,0,0,'','','','',0,'','','',0,'','','','',1,'Q',0)}";
rsDocumentation.CursorType = 0;
rsDocumentation.CursorLocation = 2;
rsDocumentation.LockType = 3;
rsDocumentation.Open();

var rsAreaCode = Server.CreateObject("ADODB.Recordset");
rsAreaCode.ActiveConnection = MM_cnnASP02_STRING;
rsAreaCode.Source = "{call dbo.cp_area_code(0,'',0,2,'Q',0)}";
rsAreaCode.CursorType = 0;
rsAreaCode.CursorLocation = 2;
rsAreaCode.LockType = 3;
rsAreaCode.Open();
%>
<html>
<head>
	<title>Update Disability Documentaion</title>
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
				document.frm0601.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (!CheckTextArea(document.frm0601.MedicalComments, 1024)){
			alert("Medical comments cannot exceed 1024 characters.");
			return ;
		}		
		if (!CheckTextArea(document.frm0601.CaseManagerComments, 1024)){
			alert("Case manager comments cannot exceed 1024 characters.");
			return ;
		}			
	
		if (!CheckDate(document.frm0601.DocumentationDate.value)) {
			alert("Invaild Documentation Date.");
			document.frm0601.DocumentationDate.focus();
			return ;
		}
		if (!CheckDate(document.frm0601.DateReceived.value)) {
			alert("Invalid Date Received.");
			document.frm0601.DateReceived.focus();
			return ;
		}

		switch(document.frm0601.Type.value){
			case "1":
				document.frm0601.TypeMedical.value="1";
			break;
			case "2":
				document.frm0601.TypeAudiology.value="1";
			break;
			case "3":
				document.frm0601.TypePsychoEd.value="1";
			break;
			default:				
			break;
		}		
		document.frm0601.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0601.Type.focus()">
<form name="frm0601" method="POST" action="<%=MM_editAction%>">
<h5>Update Disability Documentation</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="1" accesskey="F">
			<option value="1" <%=((rsDocumentation.Fields.Item("chvType").Value == "Medical")?"SELECTED":"")%>>Medical
			<option value="2" <%=((rsDocumentation.Fields.Item("chvType").Value == "Audiology")?"SELECTED":"")%>>Audiology
			<option value="3" <%=((rsDocumentation.Fields.Item("chvType").Value == "Psycho_Ed")?"SELECTED":"")%>>Psycho-Ed
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Location of Documentation:</td>
		<td nowrap><select name="LocationOfDocumentation" tabindex="2">
			<option value="A" <%=((rsDocumentation.Fields.Item("chvLoc").Value == "ASP")?"SELECTED":"")%>>ASP
			<option value="D" <%=((rsDocumentation.Fields.Item("chvLoc").Value == "DSS")?"SELECTED":"")%>>DSS
			<option value="S" <%=((rsDocumentation.Fields.Item("chvLoc").Value == "SSB")?"SELECTED":"")%>>SSB
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Documentation Date:</td>
		<td nowrap>
			<input type="text" name="DocumentationDate" value="<%=FilterDate(rsDocumentation.Fields.Item("dtsDocDate").Value)%>" size="11" maxlength="10" tabindex="3" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" value="<%=FilterDate(rsDocumentation.Fields.Item("dtsRxdate").Value)%>" size="11" maxlength="10" tabindex="4" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>Diagnosis:</td>
		<td nowrap><input type="text" name="Diagnosis" value="<%=(rsDocumentation.Fields.Item("chvDiagnosis").Value)%>" maxlength="100" size="70" tabindex="5" ></td>
    </tr>
    <tr> 
		<td nowrap>Permanent:</td>
		<td nowrap><select name="Permanent" tabindex="6">
			<option <%=((rsDocumentation.Fields.Item("bitPermanent").Value == "Yes")?"SELECTED":"")%> value="1">Yes
			<option <%=((rsDocumentation.Fields.Item("bitPermanent").Value == "No")?"SELECTED":"")%> value="0">No
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Specialist:</td>
		<td nowrap><input type="text" name="Specialist" value="<%=(rsDocumentation.Fields.Item("chvSpecialist").Value)%>" maxlength="50" size="50" tabindex="7" ></td>
    </tr>
    <tr> 
		<td nowrap>Credentials:</td>
		<td nowrap><input type="text" name="Credentials" value="<%=(rsDocumentation.Fields.Item("chvCredentials").Value)%>" maxlength="50" size="50" tabindex="8" ></td>
    </tr>
    <tr> 
		<td nowrap valign="top">Medical Comments:</td>
		<td nowrap valign="top"><textarea name="MedicalComments" rows="3" cols="65" tabindex="9"><%=(rsDocumentation.Fields.Item("chvMedical_note").Value)%></textarea></td>
    </tr>
    <tr> 
		<td nowrap>Phone Number:</td>
		<td nowrap>
			<select name="PhoneAreaCode" tabindex="10">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%=((rsDocumentation.Fields.Item("chrPh_Areacd").Value == rsAreaCode.Fields.Item("chvAC_num").Value)?"SELECTED":"")%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>			
			</select>
			<input type="text" name="PhoneNumber" value="<%=(rsDocumentation.Fields.Item("chrPh_num").Value)%>" size="9" maxlength="8" tabindex="11" onKeypress="AllowNumericOnly();" onChange="FormatPhoneNumberOnly(this);">Ext 
			<input type="text" name="PhoneExtension" value="<%=(rsDocumentation.Fields.Item("chrPh_Ext").Value)%>" size="3" maxlength="5" tabindex="12" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>Eligible for ASP:</td>
		<td nowrap><select name="EligibleForASP" tabindex="13">
			<option value="1" <%=((rsDocumentation.Fields.Item("bitEligible").Value == "Yes")?"SELECTED":"")%>>Yes
			<option value="0" <%=((rsDocumentation.Fields.Item("bitEligible").Value == "No")?"SELECTED":"")%>>No
		</select></td>
    </tr>
    <tr> 
		<td nowrap valign="top">Case Manager Comments:</td>
		<td nowrap valign="top"><textarea name="CaseManagerComments" rows="3" cols="65" tabindex="14" accesskey="L"><%=(rsDocumentation.Fields.Item("chvNote").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="16" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="17" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="TypeMedical" value="0">
<input type="hidden" name="TypeAudiology" value="0">
<input type="hidden" name="TypePsychoEd" value="0">
<input type="hidden" name="MM_action" value="update">
<input type="hidden" name="MM_recordId" value="<%=rsDocumentation.Fields.Item("intDbtyDoc_id").Value%>">
</form>
</body>
</html>
<%
rsDocumentation.Close();
rsAreaCode.Close();
%>
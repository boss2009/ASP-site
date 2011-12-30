<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

// boolean to abort record edit
var MM_abortEdit = false;

// query string to execute
var MM_editQuery = "";
%>
<%
// *** Insert Record: set variables
if (String(Request("MM_insert")) != "undefined") {
  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable  = "dbo.tbl_Ext_FS";
  var MM_editRedirectUrl = "InsertSuccessful.html";
  var MM_fieldsStr = "AdultID|value|AgencyType|value|EntryDate|value|EligibleForCSG|value|EligibleForEPPD|value|ClaimNumber|value|ClaimStatus|value|ClosingDate|value|SettlementLetterReceivedDate|value|ContactFirstName|value|ContactMiddleName|value|ContactLastName|value|ContactPhoneAreaCode|value|ContactPhoneNumber|value|ContactPhoneExtension|value|Comments|value";
  var MM_columnsStr = "intAdult_id|none,none,NULL|chrExtFS_chbx|',none,''|dtsEntry_date|',none,''|bitIsElgb4_ASP|none,none,NULL|chrElgb_VR|',none,''|chvClaim_no|',none,''|bitActive_Claim|none,none,NULL|dtsClose_date|',none,''|dtsLetter_rx|',none,''|chvCtcFst_name|',none,''|chvCtcMdl_name|',none,''|chvCtcLst_name|',none,''|chvCtcPh_Arcd|',none,''|chvCtcPh_Num|',none,''|chvCtcPh_Ext|',none,''|chvComment|',none,''";

  // create the MM_fields and MM_columns arrays
  var MM_fields = MM_fieldsStr.split("|");
  var MM_columns = MM_columnsStr.split("|");
  
  // set the form values
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    MM_fields[i+1] = String(Request.Form(MM_fields[i]));
  }

  // append the query string to the redirect URL
  if (MM_editRedirectUrl && Request.QueryString && Request.QueryString.length > 0) {
    MM_editRedirectUrl += ((MM_editRedirectUrl.indexOf('?') == -1)?"?":"&") + Request.QueryString;
  }
}
%>
<%
// *** Insert Record: construct a sql insert statement and execute it

if (String(Request("MM_insert")) != "undefined") {

  // create the sql insert statement
  var MM_tableValues = "", MM_dbValues = "";
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    var formVal = MM_fields[i+1];
    var MM_typesArray = MM_columns[i+1].split(",");
    var delim =    (MM_typesArray[0] != "none") ? MM_typesArray[0] : "";
    var altVal =   (MM_typesArray[1] != "none") ? MM_typesArray[1] : "";
    var emptyVal = (MM_typesArray[2] != "none") ? MM_typesArray[2] : "";
    if (formVal == "" || formVal == "undefined") {
      formVal = emptyVal;
    } else {
      if (altVal != "") {
        formVal = altVal;
      } else if (delim == "'") { // escape quotes
        formVal = "'" + formVal.replace(/'/g,"''") + "'";
      } else {
        formVal = delim + formVal + delim;
      }
    }
    MM_tableValues += ((i != 0) ? "," : "") + MM_columns[i];
    MM_dbValues += ((i != 0) ? "," : "") + formVal;
  }
  MM_editQuery = "insert into " + MM_editTable + " (" + MM_tableValues + ") values (" + MM_dbValues + ")";

  if (!MM_abortEdit) {
    // execute the insert
    var MM_editCmd = Server.CreateObject('ADODB.Command');
    MM_editCmd.ActiveConnection = MM_editConnection;
    MM_editCmd.CommandText = MM_editQuery;
    MM_editCmd.Execute();
    MM_editCmd.ActiveConnection.Close();

    if (MM_editRedirectUrl) {
      Response.Redirect(MM_editRedirectUrl);
    }
  }
}
%>
<%
var rsReferral = Server.CreateObject("ADODB.Recordset");
rsReferral.ActiveConnection = MM_cnnASP02_STRING;
rsReferral.Source = "{call dbo.cp_Referrals("+ Request.QueryString("intAdult_id") + ")}";
rsReferral.CursorType = 0;
rsReferral.CursorLocation = 2;
rsReferral.LockType = 3;
rsReferral.Open();

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
	<title>New External Agency</title>
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
	function Save(){
		if (!CheckDate(document.frm0605.EntryDate.value)) {
			alert("Invalid Entry Date.");
			document.frm0605.EntryDate.focus();
			return ;
		}
		if (!CheckDate(document.frm0605.ClosingDate.value)) {
			alert("Invalid Closing Date.");
			document.frm0605.ClosingDate.focus();
			return ;
		}
		if (!CheckDate(document.frm0605.SettlementLetterReceivedDate.value)) {
			alert("Invalid Settlement Letter Received Date.");
			document.frm0605.SettlementLetterReceivedDate.focus();
			return ;
		}
		document.frm0605.submit();
		document.frm0605.btnSave.disabled = true;
	}
	</script>
</head>
<body onLoad="javascript:document.frm0605.EntryDate.focus()" >
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0605">
<h5>New External Agency</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Entry Date:</td>
		<td nowrap><input type="text" name="EntryDate" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)">
    </tr>	
    <tr> 	
		<td nowrap>Referral Date:</td>
		<td nowrap><select name="ReferralDate" tabindex="2">
		<% 
		while (!rsReferral.EOF) {
		%>
			<option value="<%=(rsReferral.Fields.Item("intadult_id").Value)%>" <%=((rsReferral.Fields.Item("intadult_id").Value == Request.QueryString("intAdult_id"))?"SELECTED":"")%> ><%=FilterDate(rsReferral.Fields.Item("dtsRefral_date").Value)%>, <%=(rsReferral.Fields.Item("chvDetails").Value)%></option>
		<%
			rsReferral.MoveNext();
		}
		%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Agency Type:</td>
		<td nowrap><select name="AgencyType" tabindex="3">
			<option value="1">ICBC or other Auto Insurance
			<option value="2">WCB
			<option value="3">LTD
			<option value="4">Veteran's Affairs				  	
		</select></td> 
    </tr>
    <tr> 
		<td nowrap>Claim Number:</td>
		<td nowrap><input type="text" name="ClaimNumber" maxlength="30" tabindex="4"></td>
    </tr>
    <tr> 
		<td nowrap>Claim Status:</td>
		<td nowrap><select name="ClaimStatus" tabindex="5">
			<option value="1">Active
			<option value="0">Inactive
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Closing Date:</td>
		<td nowrap><input type="text" name="ClosingDate" size="11" maxlength="10" tabindex="6" onChange="FormatDate(this)" ></td>
	</tr>
	<tr>
		<td nowrap width="100">Settlement Letter Received:</td>
		<td nowrap>
			<input type="text" name="SettlementLetterReceivedDate" size="11" maxlength="10" tabindex="7" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>Eligible for CSG:</td>
		<td nowrap><select name="EligibleForCSG" tabindex="8">
		  	<option value="1">Yes
			<option value="0" SELECTED>No
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Eligible for EPPD:</td>
		<td nowrap><select name="EligibleForEPPD" tabindex="9">
			<option value="1" SELECTED>Yes
			<option value="2">No
			<option value="0">N/A
		</select></td>
    </tr>	
    <tr> 
		<td nowrap colspan="2">Contact Person:</td>
	<tr>
        <td nowrap align="right">First Name:</td>
		<td nowrap align="left"><input type="text" name="ContactFirstName" maxlength="50" tabindex="10"></td>
	</tr>
	<tr>
		<td nowrap align="right">Last Name:</td>
		<td nowrap align="left"><input type="text" name="ContactLastName" maxlength="50" tabindex="12"></td>
    </tr>
    <tr> 
		<td nowrap align="right">Phone number:</td>
		<td nowrap align="left"> 
			<select name="ContactPhoneAreaCode" tabindex="13">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>"><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			rsAreaCode.MoveFirst();
			%>
			</select>
			<input type="text" name="ContactPhoneNumber" maxlength="8" tabindex="14" size="9" onKeypress="AllowNumericOnly();" onChange="FormatPhoneNumberOnly(this);">
			<input type="text" name="ContactPhoneExtension" maxlength="5" tabindex="15" size="3" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap valign="top">Comments:</td>
		<td nowrap valign="top"><textarea name="Comments" cols="65" rows="3" tabindex="16" accesskey="L"></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" name="btnSave" value="Save" tabindex="17" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="18" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="AdultID" value="<%=(Request.QueryString("intadult_id"))%>">
<input type="hidden" name="ContactMiddleName" value="">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsReferral.Close();
rsAreaCode.Close();
%>
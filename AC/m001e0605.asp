<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var MM_abortEdit = false;

var MM_editQuery = "";
if (String(Request("MM_update")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable  = "dbo.tbl_Ext_FS";
  var MM_editColumn = "intExtFS_id";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "UpdateSuccessful.asp?page=m001q0605.asp&intAdult_id="+Request.QueryString("intAdult_id");
  var MM_fieldsStr = "AgencyType|value|EntryDate|value|EligibleForCSG|value|EligibleForEPPD|value|ClaimNumber|value|ClaimStatus|value|ClosingDate|value|SettlementLetterReceivedDate|value|ContactFirstName|value|ContactMiddleName|value|ContactLastName|value|ContactPhoneAreaCode|value|ContactPhoneNumber|value|ContactPhoneExtension|value|Notes|value";
  var MM_columnsStr = "chrExtFS_chbx|',none,''|dtsEntry_date|',none,''|bitIsElgb4_ASP|none,none,NULL|chrElgb_VR|',none,''|chvClaim_no|',none,''|bitActive_Claim|none,none,NULL|dtsClose_date|',none,''|dtsLetter_rx|',none,''|chvCtcFst_name|',none,''|chvCtcMdl_name|',none,''|chvCtcLst_name|',none,''|chvCtcPh_Arcd|',none,''|chvCtcPh_Num|',none,''|chvCtcPh_Ext|',none,''|chvComment|',none,''";

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
// *** Update Record: construct a sql update statement and execute it

if (String(Request("MM_update")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  // create the sql update statement
  MM_editQuery = "update " + MM_editTable + " set ";
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
    MM_editQuery += ((i != 0) ? "," : "") + MM_columns[i] + " = " + formVal;
  }
  MM_editQuery += " where " + MM_editColumn + " = " + MM_recordId;

  if (!MM_abortEdit) {
    // execute the update
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

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_Idv_Ext_FS("+ Request.QueryString("intAdult_id") + ","+ Request.QueryString("intExtFS_id") + ")}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();

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
	<title>Update External Agency</title>
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
				document.frm0605.reset();
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
		if (!CheckTextArea(document.frm0605.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
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
	}
	</script>
</head>
<body onLoad="javascript:document.frm0605.EntryDate.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0605">
<h5>Update External Agency</h5>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Entry Date:</td>
		<td nowrap>
			<input type="text" name="EntryDate" value="<%=FilterDate(rsFundingSource.Fields.Item("dtsEntry_date").Value)%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr> 
		<td nowrap>Agency Type:</td>
		<td nowrap><select name="AgencyType" tabindex="2">
	 		<option value="1" <%=((rsFundingSource.Fields.Item("chrExtFS_chbx").Value == 1)?"SELECTED":"")%>>ICBC or other Auto- Insur
			<option value="2" <%=((rsFundingSource.Fields.Item("chrExtFS_chbx").Value == 2)?"SELECTED":"")%>>WCB
			<option value="3" <%=((rsFundingSource.Fields.Item("chrExtFS_chbx").Value == 3)?"SELECTED":"")%>>LTD
			<option value="4" <%=((rsFundingSource.Fields.Item("chrExtFS_chbx").Value == 4)?"SELECTED":"")%>>Veteran's Affairs
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Claim Number:</td>
		<td nowrap><input type="text" name="ClaimNumber" value="<%=(rsFundingSource.Fields.Item("chvClaim_no").Value)%>" maxlength="30" size="10" tabindex="4" ></td>
    </tr>
    <tr> 
		<td nowrap>Claim Status:</td>
		<td nowrap><select name="ClaimStatus" tabindex="5">
			<option value="1" <%=((rsFundingSource.Fields.Item("bitActive_Claim").Value == 1)?"SELECTED":"")%>>Active
			<option value="0" <%=((rsFundingSource.Fields.Item("bitActive_Claim").Value == 0)?"SELECTED":"")%>>Inactive
		</select></td>
	</tr>
	<tr>
		<td nowrap>Closing Date:</td>
		<td nowrap>
			<input type="text" name="ClosingDate" value="<%=FilterDate(rsFundingSource.Fields.Item("dtsClose_date").Value)%>" size="11" maxlength="10" tabindex="6" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Settlement Letter Received:</td>			
		<td nowrap>
			<input type="text" name="SettlementLetterReceivedDate" value="<%=FilterDate(rsFundingSource.Fields.Item("dtsLetter_rx").Value)%>" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>Eligible for CSG:</td>
		<td nowrap><select name="EligibleForCSG" tabindex="10">
			<option value="1" <%=((rsFundingSource.Fields.Item("bitIsElgb4_ASP").Value == 1)?"SELECTED":"")%>>Yes
			<option value="0" <%=((rsFundingSource.Fields.Item("bitIsElgb4_ASP").Value == 0)?"SELECTED":"")%>>No				
		</select></td> 
    </tr>
    <tr> 
		<td nowrap>Eligible for EPPD:</td>
		<td nowrap><select name="EligibleForEPPD" tabindex="11"> 
			<option value="1" <%=((rsFundingSource.Fields.Item("chrElgb_VR").Value == 1)?"SELECTED":"")%>>Yes
			<option value="2" <%=((rsFundingSource.Fields.Item("chrElgb_VR").Value == 2)?"SELECTED":"")%>>No				
			<option vlaue="0" <%=((rsFundingSource.Fields.Item("chrElgb_VR").Value == 0)?"SELECTED":"")%>>N/A
		</select></td>
	</tr>	
    <tr> 
		<td nowrap>Contact Person:</td>
		<td nowrap></td>
	</tr>
	<tr>		
		<td nowrap align="right">First Name:</td>
		<td nowrap><input type="text" name="ContactFirstName" value="<%=(rsFundingSource.Fields.Item("chvCtcFst_name").Value)%>" maxlength="50" tabindex="12" ></td>
	</tr>
	<tr>
		<td nowrap align="right">Last Name:</td>
		<td nowrap><input type="text" name="ContactLastName" value="<%=(rsFundingSource.Fields.Item("chvCtcLst_name").Value)%>" maxlength="50" tabindex="14" ></td>
    </tr>
    <tr> 
		<td nowrap align="right">Phone Number:</td>
		<td nowrap>
			<select name="ContactPhoneAreaCode" tabindex="15">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%=((rsFundingSource.Fields.Item("chvCtcPh_Arcd").Value == rsAreaCode.Fields.Item("chvAC_num").Value)?"SELECTED":"")%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>
			</select>
			<input type="text" name="ContactPhoneNumber" value="<%=(rsFundingSource.Fields.Item("chvCtcPh_Num").Value)%>" maxlength="8" size="9" tabindex="16" onChange="FormatPhoneNumberOnly(this);">
			<input type="text" name="ContactPhoneExtension" value="<%=(rsFundingSource.Fields.Item("chvCtcPh_Ext").Value)%>" maxlength="5" size="3" tabindex="17" >
		</td>
    </tr>
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="18" accesskey="L"><%=(rsFundingSource.Fields.Item("chvComment").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="19" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="20" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="21" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="ContactMiddleName">
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsFundingSource.Fields.Item("intExtFS_id").Value %>">
</form>
</body>
</html>
<%
rsFundingSource.Close();
rsAreaCode.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
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
  var MM_editTable  = "dbo.tbl_referring_agent";
  var MM_editRedirectUrl = "AddDeleteSuccessful.asp?action=Add";
  var MM_fieldsStr = "ReferringAgentName|value|SetActive|value|IsLoan|value|IsBuyOut|value|FundingSourceCode|value";
  var MM_columnsStr = "chvname|',none,''|bitis_active|none,none,NULL|bitis_loan|none,none,NULL|bitis_BuyOut|none,none,NULL|chrFS_chbx|',none,''";

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
<html>
<head>
	<title>New Referring Agent</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
</head>
<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0312b">
<h5>New Referring Agent</h5>
To add this referring agent, click [Proceed]
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td>Referring Agent Name:</td>
		<td><input type="text" name="ReferringAgentName" value="<%=(Request.QueryString("chvname"))%>"  readonly accesskey="F"></td>
    </tr>
    <tr>
		<td>Is Active</td>
		<td><input type="text" name="SetActive" value="<%=(Request.QueryString("bitis_active"))%>"  readonly></td>
    </tr>
    <tr>
		<td>Is Loan:</td>
		<td><input type="text" name="IsLoan" value="<%=(Request.QueryString("bitis_loan"))%>"  readonly></td>
    </tr>
    <tr>
		<td>Is Buyout:</td>
		<td><input type="text" name="IsBuyOut" value="<%=(Request.QueryString("bitis_BuyOut"))%>"  readonly></td>
    </tr>
    <tr>
		<td>Funding Source:</td>
		<td><input type="text" name="FundingSourceCode" value="<%=(Request.QueryString("chrFS_chbx"))%>"  readonly accesskey="L"></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td><input type="submit" value="Proceed" class="btnstyle"></td>
		<td><input type="button" value="Cancel" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
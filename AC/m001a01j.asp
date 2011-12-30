<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
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
  var MM_editTable  = "dbo.tbl_Desktop";
  var MM_editRedirectUrl = "InsertSuccessful.html";
  var MM_fieldsStr = "insStaff_id|value|SaveTo|value|intAdult_Id|value|DateCopied|value";
  var MM_columnsStr = "insStaff_Id|none,none,NULL|insMOD_id|none,none,NULL|intObject_id|none,none,NULL|dtmLast_Open|',none,''";

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

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

var rsModule = Server.CreateObject("ADODB.Recordset");
rsModule.ActiveConnection = MM_cnnASP02_STRING;
rsModule.Source = "{call dbo.cp_ASP_Lkup(700)}";
rsModule.CursorType = 0;
rsModule.CursorLocation = 2;
rsModule.LockType = 3;
rsModule.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_Idv_Staff("+ Session("insStaff_id")+ ")}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title>Save to Desktop</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<form name="frm01j02" method="POST" action="<%=MM_editAction%>">
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Client Name:</td>
		<td><input type="text" name="ClientName" value="<%=(rsClient.Fields.Item("chvName").Value)%>" readonly size="30" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td>Date Copied:</td>
		<td>
			<input type="text" name="DateCopied" value="<%=CurrentDate()%>" readonly tabindex="3" accesskey="L">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
    </tr>
    <tr> 
		<td colspan="2" align="center">
			<input type="submit" value="Save to Desktop" class="btnstyle">
			<input type="button" value="Cancel" onClick="self.close();" class="btnstyle">
		</td>
    </tr>
</table>
<input type="hidden" name="SaveTo" value="16">
<input type="hidden" name="intAdult_Id" value="<%=(rsClient.Fields.Item("intAdult_Id").Value)%>">
<input type="hidden" name="insStaff_id" value="<%=Session("insStaff_id")%>">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsClient.Close();
rsModule.Close();
rsStaff.Close();
%>
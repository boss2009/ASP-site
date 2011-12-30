<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var MM_abortEdit = false;

var MM_editQuery = "";

if (String(Request("MM_insert")) != "undefined") {

  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable  = "dbo.tbl_pjt_mod";
  var MM_editRedirectUrl = "AddDeleteSuccessful.asp?action=Add";
  var MM_fieldsStr = "Project|value|ModuleName|value|ModuleID|value";
  var MM_columnsStr = "insPJTid|none,none,NULL|ncvMODname|',none,''|intMODno|none,none,NULL";

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

var rsProject = Server.CreateObject("ADODB.Recordset");
rsProject.ActiveConnection = MM_cnnASP02_STRING;
rsProject.Source = "SELECT *  FROM dbo.tbl_pjt_def  ORDER BY insPJTid ASC";
rsProject.CursorType = 0;
rsProject.CursorLocation = 2;
rsProject.LockType = 3;
rsProject.Open();

var rsModule = Server.CreateObject("ADODB.Recordset");
rsModule.ActiveConnection = MM_cnnASP02_STRING;
rsModule.Source = "SELECT *  FROM dbo.tbl_pjt_mod  WHERE insPJTid = 2  ORDER BY insPJTid, intMODno";
rsModule.CursorType = 0;
rsModule.CursorLocation = 2;
rsModule.LockType = 3;
rsModule.Open();
%>
<html>
<head>
	<title>New Module</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body onLoad="document.frm01.Project.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm01">
<h5>New Module</h5>
<table cellpadding="3" cellspacing="1">
	<tr> 
		<td>Project:</td>
		<td><select name="Project" tabindex="1" accesskey="F">
			<%
			while (!rsProject.EOF) {
			%>
				<option value="<%=(rsProject.Fields.Item("insPJTid").Value)%>" <%=((rsProject.Fields.Item("insPJTid").Value == rsModule.Fields.Item("insPJTid").Value)?"SELECTED":"")%> ><%=(rsProject.Fields.Item("ncvPJTName").Value)%>
			<%
				rsProject.MoveNext();
			}
			%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Module Name:</td>
		<td><input type="text" name="ModuleName" size="30" tabindex="2"></td>
    </tr>
    <tr> 
		<td nowrap>Module ID:</td>
		<td><input type="text" name="ModuleID" size="5" tabindex="3" accesskey="L"></td>
    </tr>
</table>
<table>
	<tr> 
		<td><input type="submit" value="Save" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsProject.Close();
rsModule.Close();
%>
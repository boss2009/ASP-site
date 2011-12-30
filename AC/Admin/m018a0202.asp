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

if (String(Request("MM_insert")) == "true") {

  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable  = "dbo.tbl_pjt_fctn";
  var MM_editRedirectUrl = "AddDeleteSuccessful.asp?action=Add";
  var MM_fieldsStr = "Module|value|FunctionName|value|IsItem|value|FunctionID|value|SubNumber|value";
  var MM_columnsStr = "insMODid|none,none,NULL|ncvFTNname|',none,''|bitis_Item|none,none,NULL|insFSTno|none,none,NULL|insFSTsubno|none,none,NULL";

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

var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "SELECT insFSTno, ncvFTNname, insMODid, bitis_Item FROM dbo.tbl_pjt_fctn ORDER BY insFTNid ASC";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();

var rsProject = Server.CreateObject("ADODB.Recordset");
rsProject.ActiveConnection = MM_cnnASP02_STRING;
rsProject.Source = "SELECT * FROM dbo.tbl_pjt_def ORDER BY insPJTid ASC";
rsProject.CursorType = 0;
rsProject.CursorLocation = 2;
rsProject.LockType = 3;
rsProject.Open();

var rsModule = Server.CreateObject("ADODB.Recordset");
rsModule.ActiveConnection = MM_cnnASP02_STRING;
rsModule.Source = "SELECT *  FROM dbo.tbl_pjt_mod  WHERE insPJTid = 2  ORDER BY insId ASC";
rsModule.CursorType = 0;
rsModule.CursorLocation = 2;
rsModule.LockType = 3;
rsModule.Open();
%>
<html>
<head>
	<title>New Function</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="MyFunctions.js"></script>	
</head>
<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm002">
<h5>New Function</h5>
<table>
    <tr> 
		<td>Module:</td>
		<td><select name="Module" tabindex="1" accesskey="F">
			<% 
			while (!rsModule.EOF) {
			%>
				<option value="<%=(rsModule.Fields.Item("insId").Value)%>" <%=((rsModule.Fields.Item("insId").Value == rsFunction.Fields.Item("insMODid").Value)?"SELECTED":"")%> ><%=(rsModule.Fields.Item("ncvMODname").Value)%></option>
			<%
				rsModule.MoveNext();
			}
			%>
        </select></td>
    </tr>
    <tr> 
		<td>Funstion:</td>
		<td><input type="text" name="FunctionName" tabindex="2"></td>
	</tr>
		<td>Is Item:</td>
		<td><select name="IsItem" tabindex="3"> 
		        <option value="1">Yes
				<option value="0">No
		</select></td>
    </tr>
    <tr> 
		<td>Function ID:</td>
		<td><input type="text" name="FunctionID" onKeypress="AllowNumericOnly();" tabindex="4"></td>
	</tr>
	<tr>
		<td>Sub Number:</td>
		<td><input type="text" name="SubNumber" onKeypress="AllowNumericOnly();" maxlength="2" tabindex="5" accesskey="L"></td>
    </tr>
</table>
<br>
<table>
    <tr> 
		<td><input type="submit" value="Save" class="btnstyle" tabindex="6"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsFunction.Close();
rsProject.Close();
rsModule.Close();
%>
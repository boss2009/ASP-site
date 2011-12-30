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

if (String(Request("MM_update")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable  = "dbo.tbl_pjt_Attribute";
  var MM_editColumn = "intAttribID";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "UpdateSuccessful.asp?page=m018q0203.asp";
  var MM_fieldsStr = "AttributeName|value|AttributeNumber|value|IncludeObject|value|IsLookup|value|Type|value|DesktopFile|value";
  var MM_columnsStr = "chvAttribName|',none,''|chrObjno|',none,''|bitIncludeObj|none,1,0|bitIs_lookup|none,1,0|insTypeid|none,none,NULL|chvDeskTopName|',none,''";

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

var rsType = Server.CreateObject("ADODB.Recordset");
rsType.ActiveConnection = MM_cnnASP02_STRING;
rsType.Source = "{call dbo.cp_ASP_Lkup(709)}";
rsType.CursorType = 0;
rsType.CursorLocation = 2;
rsType.LockType = 3;
rsType.Open();

var rsAttribute = Server.CreateObject("ADODB.Recordset");
rsAttribute.ActiveConnection = MM_cnnASP02_STRING;
rsAttribute.Source = "{call dbo.cp_pjt_Attribute("+ Request.QueryString("intAttribID") + ",1)}";
rsAttribute.CursorType = 0;
rsAttribute.CursorLocation = 2;
rsAttribute.LockType = 3;
rsAttribute.Open();
%>
<html>
<head>
	<title>Update Attribute</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
</head>
<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0203">
<h5>Update Attribute</h5>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Attribute Name:</td>
		<td><input type="text" name="AttributeName" value="<%=(rsAttribute.Fields.Item("chvAttribName").Value)%>" maxlength="30" tabindex="1" accesskey="F"></td>
	 </tr>
	<tr>	       
		<td>Attribute Number:</td>
		<td><input type="text" name="AttributeNumber" value="<%=(rsAttribute.Fields.Item("chrObjno").Value)%>" onKeypress="AllowNumericOnly();" maxlength="3" tabindex="2"></td>
    </tr>
    <tr> 
		<td>Include Object:</td>
        <td><input type="checkbox" name="IncludeObject" <%=((rsAttribute.Fields.Item("bitIncludeObj").Value == 1)?"CHECKED":"")%> value="checkbox" tabindex="3" class="chkstyle"></td>
	</tr>
	<tr>
		<td>Is Lookup:</td>
        <td><input type="checkbox" name="IsLookup" <%=((rsAttribute.Fields.Item("bitIs_lookup").Value == 1)?"CHECKED":"")%> value="checkbox" tabindex="4" class="chkstyle"></td>
    </tr>
    <tr> 
		<td>Type:</td>
		<td><select name="Type" tabindex="5">
			<% 
			while (!rsType.EOF) {
			%>
				<option value="<%=(rsType.Fields.Item("insTypeid").Value)%>" <%=((rsType.Fields.Item("insTypeid").Value == rsAttribute.Fields.Item("insTypeid").Value)?"SELECTED":"")%> ><%=(rsType.Fields.Item("chvTypeDesc").Value)%>
			<%
				rsType.MoveNext();
			}
			%>
		</select></td>
	</tr>
    <tr> 
		<td>Desktop File:</td>
		<td><input type="text" name="DesktopFile" value="<%=(rsAttribute.Fields.Item("chvDeskTopName").Value)%>" maxlength="50" tabindex="6" accesskey="L"></td>
    </tr>
</table>
<hr>
<table>
    <tr> 
		<td><input type="submit" value="Save" class="btnstyle" tabindex="7"></td>
		<td><input type="button" value="Close" onClick="history.go(-1);" tabindex="7" class="btnstyle"></td>	  
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsAttribute.Fields.Item("intAttribID").Value %>">
</form>
</body>
</html>
<%
rsType.Close();
rsAttribute.Close();
%>
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
  var MM_editTable  = "dbo.tbl_pjt_Attribute";
  var MM_editRedirectUrl = "AddDeleteSuccessful.asp?action=Add";
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

var rsType = Server.CreateObject("ADODB.Recordset");
rsType.ActiveConnection = MM_cnnASP02_STRING;
rsType.Source = "{call dbo.cp_ASP_Lkup(709)}";
rsType.CursorType = 0;
rsType.CursorLocation = 2;
rsType.LockType = 3;
rsType.Open();
%>
<html>
<head>
	<title>New Attribute</title>
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
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (Trim(document.frm003.AttributeName.value)=="") {
			alert("Enter Attribute Name.");
			document.frm003.AttributeName.focus();
			return ;
		}
		document.frm003.submit();
	}
	</script>	
</head>
<body onLoad="document.frm003.AttributeName.focus();">
<form name="frm003" method="POST" action="<%=MM_editAction%>">
<h5>New Attribute</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td>Attribute Name:</td>
		<td><input type="text" name="AttributeName" maxlength="50" accesskey="F" tabindex="1"></td>
	</tr>
	<tr>
		<td>Attribute Number:</td>
		<td><input type="text" name="AttributeNumber" maxlength="3" tabindex="2"></td>
    </tr>
    <tr> 
		<td>Include Object:</td>
		<td><input type="checkbox" name="IncludeObject" value="checkbox" tabindex="3" class="chkstyle"></td>
	</tr>
	<tr>
		<td>Is Lookup:</td>
		<td><input type="checkbox" name="IsLookup" value="checkbox" tabindex="4" class="chkstyle"></td>
    </tr>
    <tr> 
		<td>Type:</td>
		<td><select name="Type" tabindex="5">
			<% 
			while (!rsType.EOF) {
			%>
				<option value="<%=(rsType.Fields.Item("insTypeid").Value)%>" <%=((rsType.Fields.Item("insTypeid").Value == 1)?"SELECTED":"")%> ><%=(rsType.Fields.Item("chvTypeDesc").Value)%></option>
			<%
				rsType.MoveNext();
			}
			%>
		</select></td>
    </tr>
    <tr> 
		<td>Desktop File:</td>
		<td><input type="text" name="DesktopFile" maxlength="50" tabindex="6" accesskey="L" ></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="7" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="8" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsType.Close();
%>


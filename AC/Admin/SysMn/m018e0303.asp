<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
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
  var MM_editTable  = "dbo.tbl_Company";
  var MM_editColumn = "intCompany_id";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "m018q0303.asp";
  var MM_fieldsStr = "CompanyName|value";
  var MM_columnsStr = "chvName|',none,''";

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

var rsCompany = Server.CreateObject("ADODB.Recordset");
rsCompany.ActiveConnection = MM_cnnASP02_STRING;
rsCompany.Source = "{call dbo.cp_Company("+ Request.QueryString("intCompany_id") + ",1)}";
rsCompany.CursorType = 0;
rsCompany.CursorLocation = 2;
rsCompany.LockType = 3;
rsCompany.Open();
%>
<html>
<head>
	<title>Update Company Lookup</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0303.reset();
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
		if (Trim(document.frm0303.CompanyName.value)==""){
			alert("Enter Company Name.");
			document.frm0303.CompanyName.focus();
			return ;		
		}
		document.frm0303.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0303.CompanyName.focus();">
<form name="frm0303" method="POST" action="<%=MM_editAction%>">
<h5>Update Company Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Company Name:</td>
		<td nowrap><input type="text" name="CompanyName" value="<%=(rsCompany.Fields.Item("chvName").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsCompany.Fields.Item("intCompany_id").Value %>">
</form>
</body>
</html>
<%
rsCompany.Close();
%>

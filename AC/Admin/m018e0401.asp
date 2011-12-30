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
  var MM_editTable  = "dbo.tbl_UsrGrp";
  var MM_editColumn = "insUsrLevel";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "m018q0401.asp";
  var MM_fieldsStr = "SystemCreate|value|SystemRead|value|SystemUpdate|value|SystemDelete|value|SystemExecute|value|PasswordCreate|value|PasswordRead|value|PasswordUpdate|value|PasswordDelete|value";
  var MM_columnsStr = "bitSys_create|none,1,0|bitSys_read|none,1,0|bitSys_update|none,1,0|bitSys_delete|none,1,0|bitSys_execute|none,1,0|bitPwd_create|none,1,0|bitPwd_read|none,1,0|bitPwd_update|none,1,0|bitPwd_delete|none,1,0";

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

var rsUserGroup = Server.CreateObject("ADODB.Recordset");
rsUserGroup.ActiveConnection = MM_cnnASP02_STRING;
rsUserGroup.Source = "{call dbo.cp_Idv_UsrGrp("+ Request("insUsrLevel") + ")}";
rsUserGroup.CursorType = 0;
rsUserGroup.CursorLocation = 2;
rsUserGroup.LockType = 3;
rsUserGroup.Open();
%>
<html>
<head>
	<title>User Level Edit: <%=(rsUserGroup.Fields.Item("chvUsrLevel").Value)%></title>
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
				document.frm0401.reset();
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
		document.frm0401.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0401.SystemCreate.focus();">
<form name="frm0401" method="POST" action="<%=MM_editAction%>">
<h5>User Level Edit For <%=(rsUserGroup.Fields.Item("chvUsrLevel").Value)%></h5>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
    <tr> 
		<td><b>System Activities:</b></td>
		<td colspan="3"></td>
    </tr>
    <tr> 
		<td></td>
		<td>Create</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitSys_create").Value == 1)?"CHECKED":"")%> type="checkbox" name="SystemCreate" value="checkbox" class="chkstyle"></td>
		<td></td>
    </tr>
    <tr> 
		<td></td>
		<td>Read</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitSys_read").Value == 1)?"CHECKED":"")%> type="checkbox" name="SystemRead" value="checkbox" class="chkstyle"></td>
		<td></td>
	</tr>
	<tr> 
		<td></td>
		<td>Update</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitSys_update").Value == 1)?"CHECKED":"")%> type="checkbox" name="SystemUpdate" value="checkbox" class="chkstyle"></td>
		<td></td>
	</tr>
	<tr> 
		<td></td>
		<td>Delete</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitSys_delete").Value == 1)?"CHECKED":"")%> type="checkbox" name="SystemDelete" value="checkbox" class="chkstyle"></td>
		<td></td>
	</tr>
	<tr> 
		<td></td>
		<td>Execute</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitSys_execute").Value == 1)?"CHECKED":"")%> type="checkbox" name="SystemExecute" value="checkbox" class="chkstyle"></td>
		<td></td>
	</tr>
	<tr> 
		<td><b>Password:</b></td>
		<td colspan="3"></td>
	</tr>
	<tr> 
		<td></td>
		<td>Create</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitPwd_create").Value == 1)?"CHECKED":"")%> type="checkbox" name="PasswordCreate" value="checkbox" class="chkstyle"></td>
		<td></td>
	</tr>
	<tr> 
		<td></td>
		<td>Read</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitPwd_read").Value == 1)?"CHECKED":"")%> type="checkbox" name="PasswordRead" value="checkbox" class="chkstyle"></td>
		<td></td>
	</tr>
	<tr> 
		<td></td>
		<td>Update</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitPwd_update").Value == 1)?"CHECKED":"")%> type="checkbox" name="PasswordUpdate" value="checkbox" class="chkstyle"></td>
		<td></td>
	</tr>
	<tr> 
		<td></td>
		<td>Delete</td>
		<td><input <%=((rsUserGroup.Fields.Item("bitPwd_delete").Value == 1)?"CHECKED":"")%> type="checkbox" name="PasswordDelete" value="checkbox" class="chkstyle"></td>
		<td></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsUserGroup.Fields.Item("insUsrLevel").Value %>">
</form>
</body>
</html>
<%
rsUserGroup.Close();
%>

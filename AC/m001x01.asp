<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

// boolean to abort record edit
var MM_abortEdit = false;

// query string to execute
var MM_editQuery = "";

if (String(Request("MM_delete")) != "undefined" && String(Request("MM_recordId")) != "undefined") {

  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable = "dbo.tbl_Adult_Client";
  var MM_editColumn = "intAdult_id";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "AddDeleteSuccessful.asp?action=deleted";

  // append the query string to the redirect URL
  if (MM_editRedirectUrl && Request.QueryString && Request.QueryString.length > 0) {
    MM_editRedirectUrl += ((MM_editRedirectUrl.indexOf('?') == -1)?"?":"&") + Request.QueryString;
  }
}
%>
<%
// *** Delete Record: construct a sql delete statement and execute it

if (String(Request("MM_delete")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  // create the sql delete statement
  MM_editQuery = "delete from " + MM_editTable + " where " + MM_editColumn + " = " + MM_recordId;

  if (!MM_abortEdit) {
    // execute the delete
    var MM_editCmd = Server.CreateObject('ADODB.Command');
    MM_editCmd.ActiveConnection = MM_editConnection;
    MM_editCmd.CommandText = MM_editQuery;
    MM_editCmd.Execute();
    MM_editCmd.ActiveConnection.Close();

// Delete is succeed then redirect
    if (MM_editRedirectUrl) {
      Response.Redirect(MM_editRedirectUrl);
    }
// end if if (!MM_abortEdit)
  }

}

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

if (rsClient.EOF) Response.Redirect("AddDeleteSuccessful.asp?action=Abort");
%>
<html>
<head>
	<title>Delete <%=(rsClient.Fields.Item("chvName").Value)%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm01x01">
<h5>Delete <%=(rsClient.Fields.Item("chvName").Value)%></h5>
<hr>
<br>
<input type="submit" value="Delete" class="btnstyle">&nbsp;&nbsp;
<input type="button" value="Cancel" onClick="window.close();" class="btnstyle">
<input type="hidden" name="MM_delete" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsClient.Fields.Item("intAdult_id").Value %>">
</form>
</body>
</html>
<%
rsClient.Close();
%>
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

if (String(Request("MM_delete")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable = "dbo.tbl_LogMster";
  var MM_editColumn = "insStaff_id";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "AddDeleteSuccessful.asp?action=Delete";

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

    if (MM_editRedirectUrl) {
      Response.Redirect(MM_editRedirectUrl);
    }
  }

}

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_Idv_LogMster("+ Request.QueryString("insStaff_id") + ")}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title>Delete User: <%=(rsStaff.Fields.Item("chvName").Value)%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<form name="frm0101" method="POST" action="<%=MM_editAction%>">
<h5>Delete <%=(rsStaff.Fields.Item("chvName").Value)%></h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>User Name:</td>
		<td><%=(rsStaff.Fields.Item("chvName").Value)%></td>
    </tr>
    <tr> 
		<td>User Level:</td>
		<td><%=(rsStaff.Fields.Item("chvULDesc").Value)%></td>
    </tr>
</table>
<hr>
<table>
    <tr>
		<td><input type="submit" value="Delete" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" class="btnstyle"></td>
	</tr>
</table>

<input type="hidden" name="MM_delete" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsStaff.Fields.Item("insStaff_id").Value %>">
</form>
</body>
</html>
<%
rsStaff.Close();
%>
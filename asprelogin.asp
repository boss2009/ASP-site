<!--------------------------------------------------------------------------
* File Name: asprelogin.asp
* Title: AT-BC Login Page
* Main SP: cp_logmster
* Description: Same as asplogin.asp without graphics.  Used to renew
* session.  When login is successful, user is redirected to last visited page.
* Author: D. T.Chan
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="inc/ASPUtility.inc" -->
<!--#include file="Connections/cnnASP02.asp" -->
<%
var MM_LoginAction = Request.ServerVariables("URL");
if (Request.QueryString!="") MM_LoginAction += "?" + Request.QueryString;

var MM_valUsername=String(Request.Form("UserID"));
if (MM_valUsername != "undefined") {
  var MM_fldUserAuthorization = "insUsrLevel";
  var MM_redirectLoginSuccess;
  if (Request.Form("page") != ""){
	MM_redirectLoginSuccess=Request.QueryString("page")+"?" + Request.QueryString;
  } else {
	MM_redirectLoginSuccess="aspMenu.asp";
  }
  var MM_redirectLoginFailed="aspretry.html";
  var MM_flag="ADODB.Recordset";
  var MM_rsUser = Server.CreateObject(MM_flag);
  MM_rsUser.ActiveConnection = MM_cnnASP02_STRING;

// + Oct.10.2001
// MM_rsUser.Source = "SELECT chrUsrId, chrPwd";
  MM_rsUser.Source = "SELECT chrUsrId, chrPwd, insStaff_id";

  if (MM_fldUserAuthorization != "") MM_rsUser.Source += "," + MM_fldUserAuthorization;
  MM_rsUser.Source = "{call dbo.cp_logmster(0,'"+MM_valUsername+"','"+Request.Form("Password")+"',0,0,'V',0)}"
//  MM_rsUser.Source += " FROM dbo.tbl_LogMster WHERE chrUsrId='" + MM_valUsername + "' AND chrPwd='" + String() + "'";
  MM_rsUser.CursorType = 0;
  MM_rsUser.CursorLocation = 2;
  MM_rsUser.LockType = 3;
  MM_rsUser.Open();
  if (!MM_rsUser.EOF || !MM_rsUser.BOF) {
    // username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername;

    if (MM_fldUserAuthorization != "") {
      Session("MM_UserAuthorization") = String(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value);
// + Dec.06.2001
      Session("insStaff_id") =  MM_rsUser.Fields.Item("insStaff_id").Value ;
	  Session("TimeLoggedOn") = CurrentDateTime();
    } else {
      Session("MM_UserAuthorization") = "";
	  Session("insStaff_id") = "";
	  Session("TimeLoggedOn") = "User Not Logged On";
    }

    if (String(Request.QueryString("accessdenied")) != "undefined" && false) {
      MM_redirectLoginSuccess = Request.QueryString("accessdenied");
    }
	
    MM_rsUser.Close();
    Response.Redirect(MM_redirectLoginSuccess);
  }
  
  MM_rsUser.Close();
  Response.Redirect(MM_redirectLoginFailed);
}
%>
<html>
<head>
	<title>Authentification Requied</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="css/MyStyle.css" type="text/css">
</head>
<body onLoad="javascript:document.frmLogin.UserID.focus();">
<form name="frmLogin" method="post" action="<%=MM_LoginAction%>">
<div style="position: absolute; top: 10px; left: 10px; width: 210px; border: 1px solid #CCCCCC">
<table cellpadding="5" cellspacing="2">	
    <tr> 
		<td></td>
		<td><b>User ID:</b></td>
		<td><input type="text" name="UserID" maxlength="20" tabindex="1" value="<%=Request.Form("id")%>" ></td>
    </tr>
    <tr>
		<td></td>
		<td><b>Password:</b></td>
		<td><input type="password" name="Password" maxlength="8" tabindex="2"></td>
    </tr>
    <tr>
		<td></td>
		<td align="center" colspan="2">
			<input type="submit" value="Submit" class="btnstyle" tabindex="3">&nbsp;&nbsp;
			<input type="reset" value="Reset" tabindex="4" class="btnstyle">
		</td>
    </tr>
</table>
</div>
<input type="hidden" name="page" value="<%=Request.QueryString("page")%>">
</form>
</body>
</html>
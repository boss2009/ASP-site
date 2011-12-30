<!--------------------------------------------------------------------------
* File Name: asplogin.asp
* Title: ASP Login Page
* Main SP: cp_logmster
* Description: Main login page.  User is booted back again if login fails.
* When login is successful, user is redirected to main menu.
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="inc/ASPUtility.inc" -->
<!--#include file="Connections/cnnASP02.asp" -->



<%
var MM_LoginAction = Request.ServerVariables("URL");
if (Request.QueryString!="") MM_LoginAction += "?" + Request.QueryString;

var MM_valUsername = String(Request.Form("UserID"));

if (MM_valUsername != "undefined") {

	var MM_fldUserAuthorization="insUsrLevel";
	var MM_redirectLoginSuccess;

	if (Request.Form("page")!=""){
	
		MM_redirectLoginSuccess=Request.QueryString("page")+"?" + Request.QueryString;

	} else {
		MM_redirectLoginSuccess="aspMenu.asp";
	}

	var MM_redirectLoginFailed="aspretry.html";
	var MM_flag="ADODB.Recordset";
	var MM_rsUser = Server.CreateObject(MM_flag);
	MM_rsUser.ActiveConnection = MM_cnnASP02_STRING;
	MM_rsUser.Source = "SELECT chrUsrId, chrPwd, insStaff_id";
	if (MM_fldUserAuthorization != "") MM_rsUser.Source += "," + MM_fldUserAuthorization;
	MM_rsUser.Source = "{call dbo.cp_logmster(0,'"+MM_valUsername+"','"+Request.Form("Password")+"',0,0,'V',0)}"  
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
//			Session("MM_UserAuthorization") = 7;
//			Session("insStaff_id") =  136 ;

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
	<title>Demo Login Page</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="css/MyStyle.css" type="text/css">
    <style type="text/css">
<!--
.style1 {font-family: "Times New Roman"}
-->
    </style>
</head>
<body onLoad="javascript:document.frmLogin.UserID.focus()" background="i/bkgd.jpg">
<form name="frmLogin" method="post" action="<%=MM_LoginAction%>">
<div style="position: absolute; top: 100px; left: 180px; border: 1px solid #A1A192">
<table cellpadding="5" cellspacing="3" width="400">
	<tr> 
		<td align="center" class="GREY" style="font: 12pt; color: #A1A192"><span class="BLUE style1"><strong>Welcome to Demo Master Login</strong></span></td>
    </tr>
</table>
<br>
<table cellpadding="5" cellspacing="2">	
    <tr> 
		<td width="80"></td>
		<td><b>User ID:</b></td>
		<td><input type="text" name="UserID" maxlength="20" tabindex="1" value="<%=Request("id")%>"></td>
    </tr>
    <tr>
		<td></td>
		<td><b>Password:</b></td>
		<td><input type="password" name="Password" maxlength="8" tabindex="2" value="<%=Request("password")%>"></td>
    </tr>
    <tr>
		<td></td>
		<td align="center" colspan="2">
			<input type="submit" value="Submit" class="btnstyle" tabindex="3">&nbsp;&nbsp;
			<input type="reset" value="Reset" class="btnstyle" tabindex="4">
		</td>
    </tr>
</table>
</div>
<div style="position: absolute; font-family: verdana; top: 265; left: 210; font-size: 7pt">&copy;2005 Sirius Innovations Incorporated.  All rights reserved.</div>
<input type="hidden" name="page" value="<%=Request.QueryString("page")%>">
</form>
</body>
</html>
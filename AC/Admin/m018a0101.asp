<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.QueryString("action")) == "Add") {
	var rsAddUser = Server.CreateObject("ADODB.Recordset");
	rsAddUser.ActiveConnection = MM_cnnASP02_STRING;
	rsAddUser.Source = "{call dbo.cp_logmster("+Request.Form("Staff")+",'"+Request.Form("LoginID")+"','"+Request.Form("Password")+"',"+Request.Form("PermissionLevel")+",0,'A',0)}";
	rsAddUser.CursorType = 0;
	rsAddUser.CursorLocation = 2;
	rsAddUser.LockType = 3;
	rsAddUser.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}

var rsPermissionLevel = Server.CreateObject("ADODB.Recordset");
rsPermissionLevel.ActiveConnection = MM_cnnASP02_STRING;
rsPermissionLevel.Source = "{call dbo.cp_ASP_Lkup(701)}";
rsPermissionLevel.CursorType = 0;
rsPermissionLevel.CursorLocation = 2;
rsPermissionLevel.LockType = 3;
rsPermissionLevel.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_Lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title>New User</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				document.frm0101.submit();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>	
</head>
<body onLoad="document.frm0101.Staff.focus();">
<form name="frm0101" action="m018a0101.asp?action=Add" METHOD="POST">
<h5>New User</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Staff:</td>
		<td><select name="Staff" tabindex="1" accesskey="F">
			<% 
			while (!rsStaff.EOF) {
			%>
				<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%=((rsStaff.Fields.Item("insStaff_id").Value == 0)?"SELECTED":"")%>><%=(rsStaff.Fields.Item("chvname").Value)%></option>
			<%
				rsStaff.MoveNext();
			}
			%>
		</select></td>
	</tr>
    <tr> 
		<td nowrap>Login ID:</td>
		<td><input type="text" name="LoginID" maxlength="20" tabindex="2"></td>
    </tr>
    <tr> 
		<td nowrap>Password:</td>
		<td><input type="text" name="Password" maxlength="8" tabindex="3"></td>
    </tr>
    <tr> 
		<td nowrap>Permission Level:</td>
		<td><select name="PermissionLevel" tabindex="4" accesskey="L">
			<% 
			while (!rsPermissionLevel.EOF) {
			%>
				<option value="<%=(rsPermissionLevel.Fields.Item("insUsrLevel").Value)%>"><%=(rsPermissionLevel.Fields.Item("chvUsrLevel").Value)%></option>
			<%
				rsPermissionLevel.MoveNext();
			}
			%>
		</select></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="submit" value="Save" class="btnstyle" tabindex="5"></td>
		<td><input type="reset" value="Reset" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="7" class="btnstyle"></td>
    </tr>
</table>
</form>
</body>
</html>
<%
rsPermissionLevel.Close();
rsStaff.Close();
%>
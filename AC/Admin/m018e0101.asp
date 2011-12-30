<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) != "undefined" && String(Request("MM_recordId")) != "undefined") {
	var rsAddUser = Server.CreateObject("ADODB.Recordset");
	rsAddUser.ActiveConnection = MM_cnnASP02_STRING;
	rsAddUser.Source = "{call dbo.cp_logmster("+Request.QueryString("insStaff_id")+",'"+Request.Form("LoginID")+"','"+Request.Form("Password")+"',"+Request.Form("PermissionLevel")+",0,'E',0)}";
	rsAddUser.CursorType = 0;
	rsAddUser.CursorLocation = 2;
	rsAddUser.LockType = 3;
	rsAddUser.Open();
	Response.Redirect("m018q0101.asp");
}

var rsLogMaster = Server.CreateObject("ADODB.Recordset");
rsLogMaster.ActiveConnection = MM_cnnASP02_STRING;
rsLogMaster.Source = "{call dbo.cp_Idv_LogMster("+ Request.QueryString("insStaff_id") + ")}";
rsLogMaster.CursorType = 0;
rsLogMaster.CursorLocation = 2;
rsLogMaster.LockType = 3;
rsLogMaster.Open();

var rsPermissionLevel = Server.CreateObject("ADODB.Recordset");
rsPermissionLevel.ActiveConnection = MM_cnnASP02_STRING;
rsPermissionLevel.Source = "{call dbo.cp_ASP_Lkup(701)}";
rsPermissionLevel.CursorType = 0;
rsPermissionLevel.CursorLocation = 2;
rsPermissionLevel.LockType = 3;
rsPermissionLevel.Open();
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>Update Profile: <%=(rsLogMaster.Fields.Item("chvName").Value)%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<form name="frm0101" method="POST" action="<%=MM_editAction%>">
<h5>Update Profile: <%=(rsLogMaster.Fields.Item("chvName").Value)%></h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Login ID:</td>
		<td><input type="text" name="LoginID" value="<%=Trim(rsLogMaster.Fields.Item("chrUsrId").Value)%>" maxlength="20" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Password:</td>
		<td><input type="text" name="Password" value="<%=Trim(rsLogMaster.Fields.Item("chrPwd").Value)%>" maxlength="8" tabindex="2"></td>
    </tr>
    <tr> 
		<td nowrap>Permission Level:</td>
		<td><select name="PermissionLevel" tabindex="3" accesskey="L">
			<% 
			while (!rsPermissionLevel.EOF) {
			%>
				<option value="<%=(rsPermissionLevel.Fields.Item("insUsrLevel").Value)%>" <%=((rsPermissionLevel.Fields.Item("insUsrLevel").Value == rsLogMaster.Fields.Item("insUsrLevel").Value)?"SELECTED":"")%> ><%=(rsPermissionLevel.Fields.Item("chvUsrLevel").Value)%></option>
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
		<td><input type="submit" value="Save" class="btnstyle" tabindex="4"></td>
		<td><input type="reset" value="Reset" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m018q0101.asp';" tabindex="6" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsLogMaster.Fields.Item("insStaff_id").Value %>">
</form>
</body>
</html>
<%
rsLogMaster.Close();
rsPermissionLevel.Close();
%>
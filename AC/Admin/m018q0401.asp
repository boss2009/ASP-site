<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsUserGroup = Server.CreateObject("ADODB.Recordset");
rsUserGroup.ActiveConnection = MM_cnnASP02_STRING;
rsUserGroup.Source = "{call dbo.cp_ASP_Lkup(701)}";
rsUserGroup.CursorType = 0;
rsUserGroup.CursorLocation = 2;
rsUserGroup.LockType = 3;
rsUserGroup.Open();
%>
<html>
<head>
	<title>User Level Permissions</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<form name="frm0401" method="post" action="">
<h3>User Level Permissions</h3>
<a href="../aspMenu.asp">Master Menu</a> / <a href="m018Menu.asp">Administrative Options</a>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
    <tr> 
		<th class="headrow" nowrap align="left">User Level</th>
		<th class="headrow" nowrap align="left">Sys Create</th>
		<th class="headrow" nowrap align="left">Sys Read</th>
		<th class="headrow" nowrap align="left">Sys Update</th>
		<th class="headrow" nowrap align="left">Sys Delete</th>
		<th class="headrow" nowrap align="left">Sys Excute</th>
		<th class="headrow" nowrap align="left">Pwd Create</th>
		<th class="headrow" nowrap align="left">Pwd Read</th>
		<th class="headrow" nowrap align="left">Pwd Update</th>
		<th class="headrow" nowrap align="left">Pwd Delete</th>
    </tr>
<% 
var c = 1;
while (!rsUserGroup.EOF) { 
%>
    <tr class="<%=(((c%2)==0)?"ROWA":"ROWB")%>"> 
		<td><a href="m018e0401.asp?insUsrLevel=<%=(rsUserGroup.Fields.Item("insUsrLevel").Value)%>"><%=(rsUserGroup.Fields.Item("chvUsrLevel").Value)%></a></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvSys_create").Value)%></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvSys_read").Value)%></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvSys_update").Value)%></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvSys_delete").Value)%></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvSys_execute").Value)%></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvPwd_create").Value)%></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvPwd_read").Value)%></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvPwd_update").Value)%></td>
		<td align="center"><%=(rsUserGroup.Fields.Item("chvPwd_delete").Value)%></td>
    </tr>
<%
	c++;
	rsUserGroup.MoveNext();
}
%>
</table>
</form>
</body>
</html>
<%
rsUserGroup.Close();
%>

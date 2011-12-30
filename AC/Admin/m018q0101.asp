<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsLogMaster = Server.CreateObject("ADODB.Recordset");
rsLogMaster.ActiveConnection = MM_cnnASP02_STRING;
rsLogMaster.Source = "{call dbo.cp_logmster(0,'','',0,0,'Q',0)}";
rsLogMaster.CursorType = 0;
rsLogMaster.CursorLocation = 2;
rsLogMaster.LockType = 3;
rsLogMaster.Open();
%>
<html>
<head>
	<title>ASP Database Users</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=500,height=300,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>	
</head>
<body>
<h3>Demo Database Users</h3>
<a href="../aspMenu.asp">Master Menu</a> / <a href="m018Menu.asp">Administrative Options</a>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
    <tr> 
		<th nowrap class="headrow" align="left">Name</th>
		<th nowrap class="headrow" align="left">User Level</th>
		<th nowrap class="headrow" align="left">Login ID</th>
		<th nowrap class="headrow" align="left">Password</th>
		<th nowrap class="headrow" align="left">Email</th>		
		<th nowrap class="headrow" align="left">&nbsp;</th>
    </tr>
<% 
var c = 1;
while (!rsLogMaster.EOF) { 
%>
    <tr class="<%=(((c%2)==0)?"ROWA":"ROWB")%>"> 
		<td nowrap><a href="m018e0101.asp?insStaff_id=<%=(rsLogMaster.Fields.Item("insStaff_id").Value)%>"><%=(rsLogMaster.Fields.Item("chvName").Value)%></a></td>
		<td nowrap><%=(rsLogMaster.Fields.Item("chvULDesc").Value)%>&nbsp;</td>
		<td nowrap><%=(rsLogMaster.Fields.Item("chrUsrId").Value)%>&nbsp;</td>
		<td nowrap><%=(rsLogMaster.Fields.Item("chrPwd").Value)%>&nbsp;</td>
		<td nowrap><%=(rsLogMaster.Fields.Item("chvEmail").Value)%>&nbsp;</td>
		<td nowrap><a href="javascript: openWindow('m018x0101.asp?insStaff_id=<%=(rsLogMaster.Fields.Item("insStaff_id").Value)%>','w18x01');"><img src="../i/remove.gif" alt="Remove <%=(rsLogMaster.Fields.Item("chvName").Value)%>"></a></td>
    </tr>
<%
	c++;	
	rsLogMaster.MoveNext();
}
%>
</table>
<hr>
<table>
	<tr>
		<td><a href="javascript:openWindow('m018a0101.asp','w18A01');">Add User</a></td>
	</tr>
</table>
</body>
</html>
<%
rsLogMaster.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsSummary = Server.CreateObject("ADODB.Recordset");
rsSummary.ActiveConnection = MM_cnnASP02_STRING;
rsSummary.Source = "{call dbo.cp_pjt_statues_summary}";
rsSummary.CursorType = 0;
rsSummary.CursorLocation = 2;
rsSummary.LockType = 3;
rsSummary.Open();
%>
<html>
<head>
	<title>Maintenance Tools Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
</head>
<body>
<table align="center" cellspacing="0">
	<tr height="100">
		<td align="center"><a href="javascript: top.window.close();"><img src="../i/CA.gif" ALT="Return to Main Menu." width="68" height="50"></a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem" width="120"><a href="m020q0201.asp" target="MaintenanceToolsRightFrame">Issue Manager</a></td>
	</tr>
	<tr> 
		<td height="18px" align="center" nowrap class="MenuItem"><a href="m020s0101.asp" target="MaintenanceToolsRightFrame">Issue Search</a></td>
	</tr>
	<tr>
    	<td height="18px" align="center" nowrap class="MenuItem"><a href="m020q0101.asp" target="MaintenanceToolsRightFrame">Patron Locker</a></td>
	</tr>
	<tr>
		<td height="18px" align="center" nowrap class="MenuItem">&nbsp;</td>
	</tr>
	<tr>
		<td class="MenuItem">
			<table cellpadding="0" cellspacing="1" align="center">
			<% 
			while (!rsSummary.EOF) {
			%>
				<tr>
					<td align="right"><span style="font-size: 8pt; color: #aaaaaa; font-family: 'Arial', 'Helvetica', 'sans-serif';"><%=(rsSummary.Fields.Item("ncvStatus").Value)%>:</span></td>
					<td align="left"><span style="font-size: 8pt; color: #aaaaaa; font-family: 'Arial', 'Helvetica', 'sans-serif';"><%=(rsSummary.Fields.Item("intRecCnt").Value)%></span></td>
				</tr>
			<%
				rsSummary.MoveNext();
			}
			%>		
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
rsSummary.Close();
%>
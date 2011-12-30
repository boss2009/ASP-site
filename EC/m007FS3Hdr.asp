<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE FILE="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_FrmPanel(7)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();
%>
<html>
<head>
	<title>Equipment Class Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=400,height=200,scrollbars=1,status=1");
		return ;
	}
	</script>	
</head>
<body>
<table align="center" cellspacing="0">
	<tr>
		<td align="center"><a href="javascript: top.window.close();"><img src="../i/CA.gif" ALT="Return to Main Menu." width="68" height="50"></a></td>
	</tr>
<% 
while (!rsFunction.EOF) { 
%>
	<tr> 
		<td height="18px" class="MenuItem" align="center" width="100"><a href="<%=(rsFunction.Fields.Item("filename").Value)%>?<%=Request.QueryString%>" target="EquipmentClassFrameBody"><%=(rsFunction.Fields.Item("name").Value)%></a></td>
	</tr>
<%
	rsFunction.MoveNext();
}
%>	
	<tr>
		<td height="18px" class="MenuItem" align="center">&nbsp;</td>
	</tr>
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m007a01j.asp?<%=(Request.QueryString)%>','wj0101');" accesskey="D">Copy to DeskTop</a></td>
	</tr>	
</table>
</body>
</html>
<%
rsFunction.Close();
%>
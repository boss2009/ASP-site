<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_FrmPanel(22)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();
%>
<html>
<head>
	<title>Temp Student Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<script language="JavaScript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=400,height=200,scrollbars=1,status=1");
		return ;
	}	
	</script>
</head>
<body onLoad="window.focus();" bgcolor="#666666">
<table align="center" cellspacing="0">
	<tr>
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_student_01.jpg" ALT="Return to Main Menu." width="80" height="60"></a></div></td>
	</tr>
<% 
while (!rsFunction.EOF) { 
%>
    <tr> 
		<td height="18px" class="MenuItem" align="center" width="150px"><a href="<%=(rsFunction.Fields.Item("filename").Value)%>?intPStdnt_id=<%=Request.QueryString("intPStdnt_id")%>" target="BodyFrame"><%=(rsFunction.Fields.Item("name").Value)%></a></td>
    </tr>
<%
	rsFunction.MoveNext();
}
%>
	<tr>
		<td height="18px" class="MenuItem" align="center">&nbsp;</td>
	</tr>
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m022a01j.asp?intPStdnt_id=<%=Request.QueryString("intPStdnt_id")%>','wj2201');" accesskey="D">Copy to DeskTop</a></td>
	</tr>	
</table>
</body>
</html>
<%
rsFunction.Close();
%>
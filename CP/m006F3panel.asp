<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_FrmPanel(6)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();
%>
<html>
<head>
	<title>Organization Panel</title>
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
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_organization_02.jpg" ALT="Return to Main Menu." width="81" height="60"></a></div></td>
	</tr>
<% 
while (!rsFunction.EOF) { 
%>
    <tr> 
		<td height="18px" class="MenuItem" align="center" width="130"><a href="<%=(rsFunction.Fields.Item("filename").Value)%>?intCompany_id=<%=Request.QueryString("intCompany_id")%>" target="BodyFrame"><%=(rsFunction.Fields.Item("name").Value)%></a></td>
    </tr>
<%
	rsFunction.MoveNext();
}
%>
	<tr>
		<td height="18px" class="MenuItem" align="center">&nbsp;</td>
	</tr>
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m006a01j.asp?intCompany_id=<%=Request.QueryString("intCompany_id")%>','wj0101');" accesskey="D">Copy to DeskTop</a></td>
	</tr>	
</table>
</body>
</html>
<%
rsFunction.Close();
%>
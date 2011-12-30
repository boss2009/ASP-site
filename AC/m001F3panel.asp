<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_FrmPanel(1)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();
%>
<html>
<head>
	<title>Client Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<script language="JavaScript">
	<!--
	function MM_reloadPage(init) {  //reloads the window if Nav4 resized
		if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
			document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
		else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
	}
	MM_reloadPage(true);
	// -->
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=500,height=300,scrollbars=1,status=1");
		return ;
	}	
</script>
</head>
<body onLoad="window.focus();" bgcolor="#666666">
<form name="frmNav" method="post" action="">
<table align="center" cellspacing="0">
	<tr>
		<td align="center"><a href="javascript: top.window.close();"><img src="../i/tn_client_01.jpg" ALT="Return to Main Menu." width="81" height="60"></a></td>
	</tr>
<% 
while (!rsFunction.EOF) { 
	if (rsFunction.Fields.Item("name")!="Client History") {
%>
    <tr> 
		<td height="18px" class="MenuItem" align="center"><a href="<%=(rsFunction.Fields.Item("filename").Value)%>?intAdult_id=<%=Request.QueryString("intAdult_id")%>&ShowAll=0" target="BodyFrame"><%=(rsFunction.Fields.Item("name").Value)%></a></td>
    </tr>
<%
	}
	rsFunction.MoveNext();
}
%>
	<tr>
		<td height="18px" class="MenuItem" align="center">&nbsp;</td>
	</tr>
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m001a01j.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>','wj0101');" accesskey=D>Copy to DeskTop</a></td>
	</tr>	
	<tr> 
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m001q01j.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>','wj0101');" accesskey=I>Summary Info</a></td>
	</tr>
</table>
</form>
</body>
</html>
<%
rsFunction.Close();
%>

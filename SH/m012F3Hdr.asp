<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_FrmHdr_12("+ Request.QueryString("insSchool_id") + ")}";
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();
%>
<html>
<head>
	<title>Institution Header Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 570px; top: 10px"> 
  <table cellspacing="1" cellpadding="1">
    <tr> 
      <td><b>Temp  Referral Date:</b></td>
      <td width="200"><%=FilterDate(rsInstitution.Fields.Item("dtsRefral_date").Value)%></td>
      <td><b>Temp Status:</b></td>
      <td><%=(rsInstitution.Fields.Item("chvPILAT_Status").Value)%></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsInstitution.Close();
%>
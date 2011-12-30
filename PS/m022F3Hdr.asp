<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsPILATStudent = Server.CreateObject("ADODB.Recordset");
rsPILATStudent.ActiveConnection = MM_cnnASP02_STRING;
rsPILATStudent.Source = "{call dbo.cp_FrmHdr(22,"+ Request.QueryString("intPStdnt_id") + ")}";
rsPILATStudent.CursorType = 0;
rsPILATStudent.CursorLocation = 2;
rsPILATStudent.LockType = 3;
rsPILATStudent.Open();
%>
<html>
<head>
	<title>Temp Student Header Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 570px; top: 10px"> 
  <table cellspacing="1" cellpadding="1">
    <tr> 
      <td><b>Student Name:</b></td>
      <td width="200"><%=(rsPILATStudent.Fields.Item("chvStdent_Name").Value)%></td>
      <td><b>Case Manager:</b></td>
      <td><%=rsPILATStudent.Fields.Item("chvCase_Manager").Value%></td>
    </tr>
    <tr> 
      <td><b>Disability:</b></td>
      <td><%=rsPILATStudent.Fields.Item("chvDisability").Value%></td>
      <td><b>PEN:</b></td>
      <td><%=rsPILATStudent.Fields.Item("chrPEN_num").Value%></td>
    </tr>
    <tr> 
      <td><b>Region:</b></td>
      <td><%=rsPILATStudent.Fields.Item("chvRegion").Value%></td>
      <td></td>
      <td></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsPILATStudent.Close();
%>
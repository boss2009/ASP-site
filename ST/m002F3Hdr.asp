<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_FrmHdr(2,"+ String(Request.QueryString("insStaff_id")) + ")}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title>Staff Header Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 570px"> 
  <table cellspacing="1" cellpadding="1">
    <tr> 
      <td valign="top"><b>Region:</b></td>
      <td valign="top" width="200"><%=(rsStaff.Fields.Item("chvRegion").Value)%></td>
      <td valign="top"><b>Positions:</b></td>
      <td valign="top" width="200" colspan="3"></td>
    </tr>
    <tr> 
      <td valign="top" nowrap><b>Phone Number:</b></td>
      <td valign="top" width="200"><%=FormatPhoneNumber(rsStaff.Fields.Item("chvPhone_Type_1").Value,rsStaff.Fields.Item("chvPhone1_Arcd").Value,rsStaff.Fields.Item("chvPhone1_Num").Value,rsStaff.Fields.Item("chvPhone1_Ext").Value,rsStaff.Fields.Item("chvPhone_Type_2").Value,rsStaff.Fields.Item("chvPhone2_Arcd").Value,rsStaff.Fields.Item("chvPhone2_Num").Value,rsStaff.Fields.Item("chvPhone2_Ext").Value,"","","","")%></td>
      <td valign="top"><b>Is Active:</b></td>
      <td valign="top" width="200">Yes</td>
    </tr>
    <tr> 
      <td valign="top"><b>Job Title:</b></td>
      <td valign="top" colspan="3"><%=(rsStaff.Fields.Item("chvJobTitle").Value)%></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsStaff.Close();
%>
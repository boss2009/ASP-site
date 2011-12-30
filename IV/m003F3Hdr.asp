<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsInventoryHeader = Server.CreateObject("ADODB.Recordset");
rsInventoryHeader.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryHeader.Source = "{call dbo.cp_FrmHdr_3("+ Request.QueryString("intEquip_Set_id") + ")}";
rsInventoryHeader.CursorType = 0;
rsInventoryHeader.CursorLocation = 2;
rsInventoryHeader.LockType = 3;
rsInventoryHeader.Open();
%>
<html>
<head>
	<title>Inventory Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 580px"> 
  <table cellspacing="1" cellpadding="1">
    <tr> 
      <td valign="top" nowrap><b>Inventory Name:</b></td>
      <td valign="top" colspan="3"><%=(rsInventoryHeader.Fields.Item("chvInventory_Name").Value)%></td>
    </tr>
    <tr> 
      <td valign="top" nowrap width="120"><b>Inventory ID:</b></td>
      <td valign="top" nowrap width="130"><%=ZeroPadFormat(rsInventoryHeader.Fields.Item("intBar_Code_no").Value, 8)%></td>
      <td valign="top" nowrap><b>Current Status:</b></td>
      <td valign="top" nowrap><%=(rsInventoryHeader.Fields.Item("chvCurrent_Status").Value)%></td>
    </tr>
    <tr> 
      <%
if (rsInventoryHeader.Fields.Item("insIdvUser_id").Value != null) {
%>
      <td valign="top" nowrap><b>Individual User:</b></td>
      <td valign="top" colspan="3"><%=(rsInventoryHeader.Fields.Item("chvIdv_Usr_Nm").Value)%></td>
      <%
} else{
%>
      <td valign="top" nowrap><b>Institution User:</b></td>
      <td valign="top" colspan="3"><%=(rsInventoryHeader.Fields.Item("chvInstit_Usr_Nm").Value)%></td>
      <%
}
%>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsInventoryHeader.Close();
%>
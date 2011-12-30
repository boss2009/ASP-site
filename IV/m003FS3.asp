<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #Include File="../inc/ASPCheckLogin.inc" -->
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
<title><%=(rsInventoryHeader.Fields.Item("chvInventory_Name").Value)%> - Inventory ID: <%=ZeroPadFormat(rsInventoryHeader.Fields.Item("intBar_Code_no").Value, 8)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="140,*" frameborder="0" framespacing="0"> 
  <frame name="MenuFrame" scrolling="NO" src="m003F3panel.asp?intEquip_Set_id=<%=Request.QueryString("intEquip_Set_id")%>&intBar_Code_no=<%=Request.QueryString("intBar_Code_no")%>">
  <frameset rows="22%,78%" cols="*" resize=no frameborder="NO" border="0" framespacing="0" > 
    <frame name="HeaderFrame" scrolling="NO" resize=no  src="m003F3Hdr.asp?intEquip_Set_id=<%=Request.QueryString("intEquip_Set_id")%>&intBar_Code_no=<%=Request.QueryString("intBar_Code_no")%>">
    <frame name="BodyFrame" scrolling="YES" resize=no src="m003e0101.asp?intEquip_Set_id=<%=Request.QueryString("intEquip_Set_id")%>&intBar_Code_no=<%=Request.QueryString("intBar_Code_no")%>">
  </frameset>
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
</html>


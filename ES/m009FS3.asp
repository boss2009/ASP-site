<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!-- #Include File="../inc/ASPCheckLogin.inc" -->
<%
var rsEquipmentService = Server.CreateObject("ADODB.Recordset");
rsEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentService.Source = "{call dbo.cp_get_eqp_srv("+ Request.QueryString("intEquip_Srv_id") + ",0,0,'',1,'Q',0)}";
rsEquipmentService.CursorType = 0;
rsEquipmentService.CursorLocation = 2;
rsEquipmentService.LockType = 3;
rsEquipmentService.Open();
%>
<html>
<head>
	<title>Equipment Service No: <%=ZeroPadFormat(Request.QueryString("intEquip_Srv_id"),8)%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="147,*" frameborder="0" framespacing="0">

  <frame name="MenuFrame" scrolling="NO" src="m009F3panel.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intSrv_Dtl_id=<%=rsEquipmentService.Fields.Item("intShip_Dtl_id").Value%>">
  <frameset rows="20%,*" cols="*" frameborder="NO" border="0" framespacing="0" >
    <frame name="HeaderFrame" scrolling="NO"  src="m009F3Hdr.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>">
    <frame name="BodyFrame" scrolling="YES" src="m009FS01.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>">
  </frameset>
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
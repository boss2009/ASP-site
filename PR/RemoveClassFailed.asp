<!--------------------------------------------------------------------------
* File Name: RemoveClassFailed.asp
* Title: System Message
* Main SP: 
* Description: This page displays a message when removal of an equipment 
  class requested has failed.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>System Message</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript">
	function Init(){
		window.location.href='<%=Request.QueryString("page")%>?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>';
	}
	</script>
</head>
<body onload="Init();">
An error has occured.  Equipment class requested has not been removed.
</body>
</html>
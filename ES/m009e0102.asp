<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsIndividualUser = Server.CreateObject("ADODB.Recordset");
rsIndividualUser.ActiveConnection = MM_cnnASP02_STRING;
rsIndividualUser.Source = "{call dbo.cp_get_eqp_srv("+ Request.QueryString("intEquip_Srv_id") + ",1,0,'',8,'Q',0)}";
rsIndividualUser.CursorType = 0;
rsIndividualUser.CursorLocation = 2;
rsIndividualUser.LockType = 3;
rsIndividualUser.Open();
%>
<html>
<head>
	<title>Individual User</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</script>
</head>
<body>
<h5>Individual User</h5>
<table cellpadding="2" cellspacing="3">
	<tr> 
		<td nowrap>User Name:</td>
		<td nowrap><a href="javascript: openWindow('../AC/m001FS3.asp?intAdult_id=<%=rsIndividualUser.Fields.Item("intBuyer_id").Value%>','');"><%=rsIndividualUser.Fields.Item("chvBuyer_Name").Value%></a></td>	
	</tr>
	<tr> 
		<td nowrap>Address:</td>
		<td nowrap><%=(rsIndividualUser.Fields.Item("chvAddress").Value)%></td>
	</tr>
	<tr>
		<td nowrap>City:</td>
		<td nowrap><%=(rsIndividualUser.Fields.Item("chvCity").Value)%></td>
	</tr>
    <tr> 
		<td nowrap>Province:</td>
		<td nowrap><%=(rsIndividualUser.Fields.Item("chrprvst_abbv").Value)%></td>
	</tr>
	<tr>
		<td nowrap>Country:</td>
		<td nowrap><%=(rsIndividualUser.Fields.Item("chvcntry_name").Value)%></td>
	</tr>
	<tr>
		<td nowrap>E-mail:</td>		
	    <td nowrap><a href="mailto:<%=(rsIndividualUser.Fields.Item("chvEmail").Value)%>"><%=(rsIndividualUser.Fields.Item("chvEmail").Value)%></a></td>
    </tr>
	<tr> 
		<td nowrap>Referring Agent:</td>
		<td nowrap><%=rsIndividualUser.Fields.Item("chvReferring_Agent").Value%></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td>
			<input type="button" value="Close" tabindex="1" onClick="window.location.href='m009e0101.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>'" class="btnstyle">
		</td>
	</tr>
</table>
</body>
</html>
<%
rsIndividualUser.Close();
%>
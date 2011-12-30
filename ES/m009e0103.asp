<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsInstitutionUser = Server.CreateObject("ADODB.Recordset");
rsInstitutionUser.ActiveConnection = MM_cnnASP02_STRING;
rsInstitutionUser.Source = "{call dbo.cp_get_eqp_srv("+ Request.QueryString("intEquip_Srv_id") + ",1,0,'',9,'Q',0)}";
rsInstitutionUser.CursorType = 0;
rsInstitutionUser.CursorLocation = 2;
rsInstitutionUser.LockType = 3;
rsInstitutionUser.Open();

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_school_contacts("+ rsInstitutionUser.Fields.Item("insSchool_id").Value +",0,0,0,'Q',0)}"
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();
var ReferringAgent = "";
while (!rsContact.EOF) {
	if ((rsContact.Fields.Item("chvRelationship").Value == "ASP Referring Agent") || (rsContact.Fields.Item("chvRelationship").Value == "PILAT Referring Agent")) ReferringAgent = rsContact.Fields.Item("chvFst_Name").Value + " " + rsContact.Fields.Item("chvLst_Name").Value;
	rsContact.MoveNext();
}
%>
<html>
<head>
	<title>Institution User</title>
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
<h5>Institution User</h5>
<table cellpadding="2" cellspacing="3">
	<tr> 
		<td nowrap>Institution Name:</td>
		<td nowrap><a href="javascript: openWindow('../SH/m012FS3.asp?insSchool_id=<%=rsInstitutionUser.Fields.Item("insSchool_id").Value%>','');"><%=rsInstitutionUser.Fields.Item("chvSchool_Name").Value%></a></td>			
	</tr>
	<tr> 
		<td nowrap>Address:</td>
		<td><%=(rsInstitutionUser.Fields.Item("chvAddress").Value)%></td>
	</tr>
	<tr>
		<td nowrap>City:</td>
		<td nowrap><%=(rsInstitutionUser.Fields.Item("chvCity").Value)%></td>
	</tr>
    <tr> 
		<td nowrap>Province:</td>
		<td nowrap><%=(rsInstitutionUser.Fields.Item("chvProvince").Value)%></td>
	</tr>
	<tr>
		<td nowrap>Country:</td>
		<td nowrap><%=(rsInstitutionUser.Fields.Item("chvcntry_name").Value)%></td>
	</tr>
	<tr> 
		<td nowrap>Referring Agent:</td>
		<td nowrap><%=ReferringAgent%></td>
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
rsInstitutionUser.Close();
%>
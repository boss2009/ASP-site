<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
if (Request.QueryString("insEq_user_type")=="3") {
	var rsContact = Server.CreateObject("ADODB.Recordset");
	rsContact.ActiveConnection = MM_cnnASP02_STRING;
	rsContact.Source = "{call dbo.cp_ClnCtact2("+ Request.QueryString("intEq_user_id") + ",0,0,0,2,'Q',0)}";
	rsContact.CursorType = 0;
	rsContact.CursorLocation = 2;
	rsContact.LockType = 3;
	rsContact.Open();
} else {
	var rsContact = Server.CreateObject("ADODB.Recordset");
	rsContact.ActiveConnection = MM_cnnASP02_STRING;
	rsContact.Source = "{call dbo.cp_school_contacts("+ Request.QueryString("insInst_User_id") + ",0,0,2,'Q',0)}";
	rsContact.CursorType = 0;
	rsContact.CursorLocation = 2;
	rsContact.LockType = 3;
	rsContact.Open();
}
%>
<html>
<head>
	<title>Referring Agent</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</Script>	
</head>
<body>
<h5>Referring Agent</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 
		<th class="headrow" nowrap align="left" width="180px">Name</th>	
		<th class="headrow" nowrap align="left">Job Title</th>
		<th class="headrow" nowrap align="left">Work Type</th>
		<th class="headrow" nowrap align="left">Organization</th>
    </tr>
<% 
while (!rsContact.EOF) { 
	if (Request.QueryString("insEq_user_type")=="3"){
%>
    <tr> 
		<td nowrap><a href="javascript: openWindow('../CT/m004FS3.asp?intContact_id=<%=(rsContact.Fields.Item("intContact_id").Value)%>');"><%=(rsContact.Fields.Item("chvContact_Name").Value)%></a>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("chvJob_title").Value)%>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("chvWork_Type").Value)%>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("Inst_Company").Value)%>&nbsp;</td>
    </tr>
<%
	} else {
%>
    <tr> 
		<td nowrap><a href="javascript: openWindow('../CT/m004FS3.asp?intContact_id=<%=(rsContact.Fields.Item("intContact_id").Value)%>');"><%=(rsContact.Fields.Item("chvLst_Name").Value)%>, <%=(rsContact.Fields.Item("chvFst_Name").Value)%></a>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("chvJob_Title").Value)%>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("chvWork_type_desc").Value)%>&nbsp;</td>
		<td nowrap>&nbsp;</td>
    </tr>
<%
	}
	rsContact.MoveNext();
}
%>
</table>
<br><br><br>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>
<%
rsContact.Close();
%>
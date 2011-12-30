<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsCompany = Server.CreateObject("ADODB.Recordset");
rsCompany.ActiveConnection = MM_cnnASP02_STRING;
rsCompany.Source = "{call dbo.cp_Company2("+Request.QueryString("intCompany_id")+",'',0,0,0,0,0,1,0,'',1,'Q',0)}"
rsCompany.CursorType = 0;
rsCompany.CursorLocation = 2;
rsCompany.LockType = 3;
rsCompany.Open();	

var rsCompanyContact = Server.CreateObject("ADODB.Recordset");
rsCompanyContact.ActiveConnection = MM_cnnASP02_STRING;
rsCompanyContact.Source = "{call dbo.cp_Company_Contact("+ Request.QueryString("intCompany_id") + ","+rsCompany.Fields.Item("insWork_Typ_id").Value+",0,'Q',0)}";
rsCompanyContact.CursorType = 0;
rsCompanyContact.CursorLocation = 2;
rsCompanyContact.LockType = 3;
rsCompanyContact.Open();
%>
<html>
<head>
	<title>Company Contacts</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=670,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</Script>	
</head>
<body>
<h5>Organization Contacts</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 
		<th class="headrow" nowrap align="left" width="180">Name</th>	
		<th class="headrow" nowrap align="left" width="160">Job Title</th>
		<th class="headrow" nowrap align="left" width="100">Phone Number</th>
		<th class="headrow" nowrap align="left" width="120">EMail</th>
		<th class="headrow" nowrap align="left">&nbsp;</th>
    </tr>
<% 
while (!rsCompanyContact.EOF) { 
	var obj = rsCompanyContact.Fields;
%>
    <tr> 
		<td nowrap><a href="javascript: openWindow('../CT/m004FS3.asp?intContact_id=<%=(obj.Item("intContact_id").Value)%>&intCompany_id=<%=Request.QueryString("intCompany_id")%>');"><%=(obj.Item("chvLst_Name").Value)%>, <%=(obj.Item("chvFst_Name").Value)%></a>&nbsp;</td>
		<td nowrap><%=(obj.Item("chvJob_title").Value)%>&nbsp;</td>
		<td nowrap><%=FormatPhoneNumber(obj.Item("chvPhone_Type_1").Value,obj.Item("chvPhone1_Arcd").Value,obj.Item("chvPhone1_Num").Value,obj.Item("chvPhone1_Ext").Value,obj.Item("chvPhone_Type_2").Value,obj.Item("chvPhone2_Arcd").Value,obj.Item("chvPhone2_Num").Value,obj.Item("chvPhone2_Ext").Value,obj.Item("chvPhone_Type_3").Value,obj.Item("chvPhone3_Arcd").Value,obj.Item("chvPhone3_Num").Value,obj.Item("chvPhone3_Ext").Value,obj.Item("chvPhone3_Ext").Value)%>&nbsp;</td>
		<td nowrap><%=(obj.Item("chvEmail").Value)%>&nbsp;</td>
		<td nowrap><a href="javascript: openWindow('m006q0301x.asp?intCompany_id=<%=Request.QueryString("intCompany_id")%>&intContact_id=<%=(obj.Item("intContact_id").Value)%>');"><img src="../i/remove.gif" ALT="Remove <%=(obj.Item("chvLst_Name").Value)%>, <%=(obj.Item("chvFst_Name").Value)%>"></a></td>
    </tr>
<%
	rsCompanyContact.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><a href="javascript: openWindow('../CT/m004a0101.asp?LinkToClass=2&LinkToObject=<%=Request.QueryString("intCompany_id")%>&WorkType=<%=rsCompany.Fields.Item("insWork_Typ_id").Value%>','winAdd');">Add Contact</a></td>
	</tr>
</table>
</body>
</html>
<%
rsCompanyContact.Close();
%>
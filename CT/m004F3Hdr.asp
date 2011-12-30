<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsContactHeader = Server.CreateObject("ADODB.Recordset");
rsContactHeader.ActiveConnection = MM_cnnASP02_STRING;
rsContactHeader.Source = "{call dbo.cp_FrmHdr(4,"+ Request.QueryString("intContact_id") + ")}";
rsContactHeader.CursorType = 0;
rsContactHeader.CursorLocation = 2;
rsContactHeader.LockType = 3;
rsContactHeader.Open();

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_contacts("+Request.QueryString("intContact_id")+",0,'','','',0,0,0,1,0,'',1,'Q',0)}"
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();	

var WorkPlace = "";
switch (String(rsContact.Fields.Item("intWork_type_id").Value)) {
	case "12":
		var rsWorkLocation = Server.CreateObject("ADODB.Recordset");
		rsWorkLocation.ActiveConnection = MM_cnnASP02_STRING;
		rsWorkLocation.Source = "select chvName as WorkLocation from tbl_school where insSchool_id = " + rsContact.Fields.Item("insWork_id").Value;
		rsWorkLocation.CursorType = 0;
		rsWorkLocation.CursorLocation = 2;
		rsWorkLocation.LockType = 3;
		rsWorkLocation.Open();
	break;
	default :
		var rsWorkLocation = Server.CreateObject("ADODB.Recordset");
		rsWorkLocation.ActiveConnection = MM_cnnASP02_STRING;
		rsWorkLocation.Source = "select chvName as WorkLocation from tbl_company where intCompany_id = " + rsContact.Fields.Item("insWork_id").Value;
		rsWorkLocation.CursorType = 0;
		rsWorkLocation.CursorLocation = 2;
		rsWorkLocation.LockType = 3;
		rsWorkLocation.Open();
	break;
}
if (!rsWorkLocation.EOF) {
	WorkPlace = rsWorkLocation.Fields.Item("WorkLocation").Value;
}
%>																									
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>Contact Header Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 570px"> 
<table cellspacing="1" cellpadding="1">
	<tr> 
		<td valign="top" nowrap><b>Name:</b></td>
		<td valign="top" width="225"><%=Trim(rsContactHeader.Fields.Item("chvFst_Name").Value)%>&nbsp;<%=Trim(rsContactHeader.Fields.Item("chvLst_Name").Value)%></td>
		<td valign="top" nowrap><b>Phone Number:</b></td>
		<td valign="top"><%=FormatPhoneNumber(rsContact.Fields.Item("chvPhone_Type_1").Value,rsContact.Fields.Item("chvPhone1_Arcd").Value,rsContact.Fields.Item("chvPhone1_Num").Value,rsContact.Fields.Item("chvPhone1_Ext").Value,"","","","","","","","")%></td>
	</tr>
	<tr> 
		<td valign="top" nowrap><b>Job Title:</b></td>
		<td valign="top"><%=Trim(rsContact.Fields.Item("chvJob_Title").Value)%></td>
		<td valign="top" nowrap><b>Work Place:</b></td>
		<td valign="top"><%=Trim(WorkPlace)%></td>
	</tr>
</table>
</div>
</body>
</html>
<%
rsContactHeader.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsContactAddress = Server.CreateObject("ADODB.Recordset");
rsContactAddress.ActiveConnection = MM_cnnASP02_STRING;
rsContactAddress.Source = "{call dbo.cp_Contact_Address("+ Request.QueryString("intContact_id") + ",0,'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,0,'Q',0)}";
rsContactAddress.CursorType = 0;
rsContactAddress.CursorLocation = 2;
rsContactAddress.LockType = 3;
rsContactAddress.Open();

var ChkContactAddress = Server.CreateObject("ADODB.Command");
ChkContactAddress.ActiveConnection = MM_cnnASP02_STRING;
ChkContactAddress.CommandText = "dbo.cp_Chk_Contact_Address";
ChkContactAddress.CommandType = 4;
ChkContactAddress.CommandTimeout = 0;
ChkContactAddress.Prepared = true;
ChkContactAddress.Parameters.Append(ChkContactAddress.CreateParameter("RETURN_VALUE", 3, 4));
ChkContactAddress.Parameters.Append(ChkContactAddress.CreateParameter("@intContact_id", 3, 1,10000,Request.QueryString("intContact_id")));
ChkContactAddress.Parameters.Append(ChkContactAddress.CreateParameter("@insRtnFlag", 2, 2));
ChkContactAddress.Execute();
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
	<title>Contact Address</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();	   
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=420,height=400,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</script>
</head>
<body>
<h5>Contact Address</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr>
		<th class="headrow" align="left" nowrap>Type</th>
		<th class="headrow" align="left" nowrap width="150">Address</th>
		<th class="headrow" align="left" nowrap width="100">City</th>
		<th class="headrow" align="left" nowrap>Province</th>
		<th class="headrow" align="left" nowrap>Country</th>
		<th class="headrow" align="left" nowrap>Postal Code</th>
		<th class="headrow" align="left" nowrap>Phone Number</th>
		<th class="headrow" align="left" nowrap>E-mail</th>
	</tr>
<% 
while (!rsContactAddress.EOF) { 
%>
    <tr> 
		<td valign="top"><a href="m004e0201.asp?intaddr_id=<%=(rsContactAddress.Fields.Item("intaddr_id").Value)%>&intContact_id=<%=Request.QueryString("intContact_id")%>"><%=((rsContactAddress.Fields.Item("chvAddrs_type").Value=="H")?"Home":"Work")%></a></td>
		<td valign="top"><%=Trim(rsContactAddress.Fields.Item("chvAddress").Value)%>&nbsp;</td>
		<td valign="top"><%=(rsContactAddress.Fields.Item("chvCity").Value)%>&nbsp;</td>
		<td valign="top" align="center"><%=(rsContactAddress.Fields.Item("chvProv").Value)%>&nbsp;</td>
		<td valign="top" align="center"><%=(rsContactAddress.Fields.Item("chvCountry").Value)%>&nbsp;</td>		
		<td valign="top" align="center"><%=FormatPostalCode(rsContactAddress.Fields.Item("chvPostal_zip").Value)%>&nbsp;</td>
		<td valign="top" nowrap><%=FormatPhoneNumber(rsContactAddress.Fields.Item("chvPhone_Type1").Value,rsContactAddress.Fields.Item("chvPhone1_Arcd").Value,rsContactAddress.Fields.Item("chvPhone1_Num").Value,rsContactAddress.Fields.Item("chvPhone1_Ext").Value,rsContactAddress.Fields.Item("chvPhone_Type2").Value,rsContactAddress.Fields.Item("chvPhone2_Arcd").Value,rsContactAddress.Fields.Item("chvPhone2_Num").Value,rsContactAddress.Fields.Item("chvPhone2_Ext").Value,"","","","")%>&nbsp;</td>
		<td valign="top"><%=(rsContactAddress.Fields.Item("chvemail").Value)%>&nbsp;</td>
    </tr>
<%
	rsContactAddress.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
<%
switch (ChkContactAddress.Parameters.Item("@insRtnFlag").Value) {
	case 0:
%>
		<td width="200"><a href="javascript: openWindow('m004a0201.asp?AddressType=Home&intContact_id=<%=Request.QueryString("intContact_id")%>','wa0201');">Add Home Address</a></td>
		<td width="200"><a href="javascript: openWindow('m004a0201.asp?AddressType=Work&intContact_id=<%=Request.QueryString("intContact_id")%>','wa0202');">Add Work Address</a></td>
<%
	break ;
	case 1:
%>
		<td width="200"><a href="javascript: openWindow('m004a0201.asp?AddressType=Work&intContact_id=<%=Request.QueryString("intContact_id")%>','wa0202');">Add Work Address</a></td>
<%
	break ;
	case 2:
%>
		<td width="200"><a href="javascript: openWindow('m004a0201.asp?AddressType=Home&intContact_id=<%=Request.QueryString("intContact_id")%>','wa0201');">Add Home Address</a></td>
<%
	break ;
}
%>
	</tr>
</table>
</body>
</html>
<%
rsContactAddress.Close();
%>
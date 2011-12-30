<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_ClnCtact2("+ Request.QueryString("intAdult_id") + ",0,0,0,0,'Q',0)}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();
%>
<html>
<head>
	<title>Referring Agent</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Referring Agent</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 
		<th class="headrow" nowrap align="left" width="180">Name</th>	
		<th class="headrow" nowrap align="left">Relationship</th>
		<th class="headrow" nowrap align="left">Phone Number</th>
		<th class="headrow" nowrap align="left">Organization</th>		
		<th class="headrow" nowrap align="left">Job Title</th>		
		<th class="headrow" nowrap align="left">Organization Type</th>
    </tr>
<%
function Lookup(id){
	var temp = "";
	switch (String(id)) {
		case "0":
			temp = "None";
		break;
		case "1":
			temp = "Home";
		break;
		case "2":
			temp = "Off";
		break
		case "3":
			temp = "Cell";
		break;
		case "4":
			temp = "Pger";
		break;
		case "5":
			temp = "Fax";
		break;
		case "6":
			temp = "Wk";
		break;
	}
	return temp;
}

while (!rsContact.EOF) { 
	if (Trim(rsContact.Fields.Item("chvRelationship").Value)=="Referring Agent") {
%>
    <tr> 
		<td nowrap><%=(rsContact.Fields.Item("chvContact_Name").Value)%>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("chvRelationship").Value)%>&nbsp;</td>
<!-- + Oct.28.2005
		<td nowrap><%=FormatPhoneNumber(Lookup(rsContact.Fields.Item("intPhone_Type_1").Value),rsContact.Fields.Item("chvPhone1_Arcd").Value,rsContact.Fields.Item("chvPhone1_Num").Value,rsContact.Fields.Item("chvPhone1_Ext").Value,Lookup(rsContact.Fields.Item("intPhone_Type_2").Value),rsContact.Fields.Item("chvPhone2_Arcd").Value,rsContact.Fields.Item("chvPhone2_Num").Value,rsContact.Fields.Item("chvPhone2_Ext").Value,"","","","")%>&nbsp;</td>				
-->
		<td nowrap><%=(rsContact.Fields.Item("Inst_Company").Value)%>&nbsp;</td>
		<td nowrap><%=(rsContact.Fields.Item("chvJob_title").Value)%>&nbsp;</td>				
		<td nowrap><%=(rsContact.Fields.Item("chvWork_Type").Value)%>&nbsp;</td>
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
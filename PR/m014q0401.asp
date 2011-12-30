<!--------------------------------------------------------------------------
* File Name: m014q0401.asp
* Title: Requested and Received Notes
* Main SP: cp_PR_BackOrder_Rx
* Description: This page lists the notes of a purchase requisition.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_Purchase_Requisition_Note("+ Request.QueryString("insPurchase_Req_id") + ",'',0,0,'',0,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();

var ChkNotes = Server.CreateObject("ADODB.Command");
ChkNotes.ActiveConnection = MM_cnnASP02_STRING;
ChkNotes.CommandText = "dbo.cp_Chk_Purchase_Requisition_Note";
ChkNotes.CommandType = 4;
ChkNotes.CommandTimeout = 0;
ChkNotes.Prepared = true;
ChkNotes.Parameters.Append(ChkNotes.CreateParameter("RETURN_VALUE", 3, 4));
ChkNotes.Parameters.Append(ChkNotes.CreateParameter("@insPR_id", 3, 1,10000,Request.QueryString("insPurchase_Req_id")));
ChkNotes.Parameters.Append(ChkNotes.CreateParameter("@insRtnFlag", 2, 2));
ChkNotes.Execute();
%>
<html>
<head>
	<title>Requested And Received Notes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();	   
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=700,height=400,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	</script>
</head>
<body>
<h5>Requested And Received Notes</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable" width="100%">
	<tr> 		
		<th class="headrow" align="left" width="120">Note Type</th>
		<th class="headrow" align="left">Description</th>	  
	</tr>
<% 
while ((!rsNotes.EOF) && (ChkNotes.Parameters.Item("@insRtnFlag").Value>0)) { 
%>
    <tr> 
		<td><a href="m014e0401.asp?int_Note_id=<%=(rsNotes.Fields.Item("int_Note_id").Value)%>&insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>"><%=(rsNotes.Fields.Item("chrNT_Desc").Value)%> Notes</a></td>		
		<td><%=(rsNotes.Fields.Item("chvNote_Desc").Value)%></td>
    </tr>
<%
	rsNotes.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
<%
if (!rsNotes.EOF) rsNotes.MoveFirst();
switch (ChkNotes.Parameters.Item("@insRtnFlag").Value) {
	case 0:
%>
		<td width="200"><a href="javascript: openWindow('m014a0401.asp?NotesType=Requested&insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','wA0401');">Add Requested Notes</a></td>
		<td><a href="javascript: openWindow('m014a0401.asp?NotesType=Received&insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','wA0401');">Add Received Notes</a></td>
<%
	break ;
	case 1:
%>
		<td><a href="javascript: openWindow('m014a0401.asp?NotesType=Received&insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','wA0401');">Add Received Notes</a></td>
<%
	break ;
	case 2:
%>
		<td><a href="javascript: openWindow('m014a0401.asp?NotesType=Requested&insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','wA0401');">Add Requested Notes</a></td>
<%
	break ;
}
%>
	</tr>
</table>
</body>
</html>
<%
rsNotes.Close();
%>
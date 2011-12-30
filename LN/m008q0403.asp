<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_eqp_loaned_notes("+ Request.QueryString("intLoan_req_id") + ",0,'',0,'',0,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();

var ChkNotes = Server.CreateObject("ADODB.Command");
ChkNotes.ActiveConnection = MM_cnnASP02_STRING;
ChkNotes.CommandText = "dbo.cp_Chk_Loan_Note";
ChkNotes.CommandType = 4;
ChkNotes.CommandTimeout = 0;
ChkNotes.Prepared = true;
ChkNotes.Parameters.Append(ChkNotes.CreateParameter("RETURN_VALUE", 3, 4));
ChkNotes.Parameters.Append(ChkNotes.CreateParameter("@insPR_id", 3, 1,10000,Request.QueryString("intLoan_req_id")));
ChkNotes.Parameters.Append(ChkNotes.CreateParameter("@insRtnFlag", 2, 2));
ChkNotes.Execute();
%>
<html>
<head>
	<title>Equipment Loaned Notes</title>
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
<h5>Equipment Loaned Notes</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 		
		<th class="headrow" align="left" nowrap valign="top">Note Type</th>
		<th class="headrow" align="left" width="400">Description</th>	  
	</tr>
<% 
while ((!rsNotes.EOF) && ((ChkNotes.Parameters.Item("@insRtnFlag").Value==1)||(ChkNotes.Parameters.Item("@insRtnFlag").Value==3))) { 
%>
    <tr> 
		<td valign="top"><a href="m008e0403.asp?int_Note_id=<%=(rsNotes.Fields.Item("int_Note_id").Value)%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>"><%=(rsNotes.Fields.Item("chvType_of_Note").Value)%></a></td>
        <td valign="top"><%=(rsNotes.Fields.Item("chvNote_Desc").Value)%></td>
    </tr>
<%
	rsNotes.MoveNext();
}
%>
</table>
<hr>
<%
if (!rsNotes.EOF) rsNotes.MoveFirst();
switch (ChkNotes.Parameters.Item("@insRtnFlag").Value) {
	case 0:
%>
<a href="javascript: openWindow('m008a0403.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>','wA0403');">Add Loaned Notes</a>
<%
	break ;
	case 2:
%>
<a href="javascript: openWindow('m008a0403.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>','wA0403');">Add Loaned Notes</a>
<%
	break ;
}
%>
</body>
</html>
<%
rsNotes.Close();
%>
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsNotesServices = Server.CreateObject("ADODB.Recordset");
rsNotesServices.ActiveConnection = MM_cnnASP02_STRING;
rsNotesServices.Source = "{call dbo.cp_query_all_pilat_srv_note("+Request.QueryString("insSchool_id")+",0)}"
rsNotesServices.CursorType = 0;
rsNotesServices.CursorLocation = 2;
rsNotesServices.LockType = 3;
rsNotesServices.Open();

var count = 0;
while (!rsNotesServices.EOF) {
	count++;
	rsNotesServices.MoveNext();
}
rsNotesServices.Requery();

var NotesServicesArray = new Array(count);

for (var i = 0; i < count; i++) {
	NotesServicesArray[i] = new Array(12);	
	NotesServicesArray[i][0] = (((rsNotesServices.Fields.item("dtsService_Date").Value==null)?rsNotesServices.Fields.item("dtsRequest_Date").Value:rsNotesServices.Fields.item("dtsService_Date").Value));
	NotesServicesArray[i][1] = rsNotesServices.Fields.item("intSrv_Note_id").Value;
	NotesServicesArray[i][2] = rsNotesServices.Fields.item("intSchool_note_id").Value;	
	NotesServicesArray[i][3] = rsNotesServices.Fields.item("intService_Req_id").Value;	
	NotesServicesArray[i][4] = rsNotesServices.Fields.item("intYearCycle").Value;	
	NotesServicesArray[i][5] = rsNotesServices.Fields.item("intYearCycle").Value;	
	NotesServicesArray[i][6] = rsNotesServices.Fields.item("chrServiceProvider").Value;	
	NotesServicesArray[i][7] = rsNotesServices.Fields.item("chvSrv_Code").Value;	
	NotesServicesArray[i][8] = rsNotesServices.Fields.item("chvNote_type").Value;	
	NotesServicesArray[i][9] = rsNotesServices.Fields.item("chvFunding_Src").Value;	
	NotesServicesArray[i][10] = rsNotesServices.Fields.item("chvNotes").Value;	
	NotesServicesArray[i][11] = ((rsNotesServices.Fields.item("dtsService_Date").Value==null)?"Service":"Notes")
	rsNotesServices.MoveNext();	
}

for (var i = count - 2; i > 0; i--){
	for (var j = 0; j <= i; j++) {
		if (NotesServicesArray[j][0] < NotesServicesArray[j+1][0]) {
			var temp = new Array(12);
			temp[0] = NotesServicesArray[j][0];
			temp[1] = NotesServicesArray[j][1];
			temp[2] = NotesServicesArray[j][2];
			temp[3] = NotesServicesArray[j][3];
			temp[4] = NotesServicesArray[j][4];
			temp[5] = NotesServicesArray[j][5];
			temp[6] = NotesServicesArray[j][6];
			temp[7] = NotesServicesArray[j][7];
			temp[8] = NotesServicesArray[j][8];
			temp[9] = NotesServicesArray[j][9];
			temp[10] = NotesServicesArray[j][10];
			temp[11] = NotesServicesArray[j][11];
			NotesServicesArray[j][0] = NotesServicesArray[j+1][0];
			NotesServicesArray[j][1] = NotesServicesArray[j+1][1];
			NotesServicesArray[j][2] = NotesServicesArray[j+1][2];
			NotesServicesArray[j][3] = NotesServicesArray[j+1][3];
			NotesServicesArray[j][4] = NotesServicesArray[j+1][4];
			NotesServicesArray[j][5] = NotesServicesArray[j+1][5];
			NotesServicesArray[j][6] = NotesServicesArray[j+1][6];
			NotesServicesArray[j][7] = NotesServicesArray[j+1][7];
			NotesServicesArray[j][8] = NotesServicesArray[j+1][8];
			NotesServicesArray[j][9] = NotesServicesArray[j+1][9];
			NotesServicesArray[j][10] = NotesServicesArray[j+1][10];
			NotesServicesArray[j][11] = NotesServicesArray[j+1][11];
			NotesServicesArray[j+1][0] = temp[0];			
			NotesServicesArray[j+1][1] = temp[1];			
			NotesServicesArray[j+1][2] = temp[2];			
			NotesServicesArray[j+1][3] = temp[3];			
			NotesServicesArray[j+1][4] = temp[4];			
			NotesServicesArray[j+1][5] = temp[5];			
			NotesServicesArray[j+1][6] = temp[6];			
			NotesServicesArray[j+1][7] = temp[7];			
			NotesServicesArray[j+1][8] = temp[8];			
			NotesServicesArray[j+1][9] = temp[9];			
			NotesServicesArray[j+1][10] = temp[10];			
			NotesServicesArray[j+1][11] = temp[11];			
		}
	} 
}
%>
<html>
<head>
	<title>Services and Notes History</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=700,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}		
	</Script>	
</head>
<body>
<h5>Services and Notes History</h5>
<table cellspacing="1">
	<tr>	
		<td width="160">Displaying All of <%=count%> Records.</td>
		<td width="130"><a href="javascript: openWindow('m012a0501.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>','wA0501');">Add Service Request</a></td>
		<td width="90"><a href="javascript: openWindow('m012a0502.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>','wA0502');">Add Notes</a></td>
<%
if (Request.QueryString("ShowFS")=="1") {
%>		
		<td nowrap><a href="m012q0501.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>&ShowFS=0">Hide Funding Source</a></td>
<%
} else {
%>
		<td nowrap><a href="m012q0501.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>&ShowFS=1">Show Funding Source</a></td>
<%
}
%>		
	</tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr>
		<th class="headrow" nowrap align="left">Date</th>
		<th class="headrow" nowrap align="left" width="190">Type</th>
		<th class="headrow" nowrap align="left">Service Provider</th>
<%		
if (Request.QueryString("ShowFS")=="1") {
%>		
		<th class="headrow" nowrap align="left">Funding Source</th>
<%
}
%>		
		<th class="headrow" nowrap align="left">Notes</th>
	</tr>
<%
for (var i = 0; i < count; i++){
	if (NotesServicesArray[i][11]=="Notes"){
%>
    <tr>
		<td valign="top"><a href="m012e0502.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intSrv_Note_id=<%=NotesServicesArray[i][1]%>&intStd_note_id=<%=NotesServicesArray[i][2]%>"><%=FilterDate(NotesServicesArray[i][0])%></a>&nbsp;</td>
		<td valign="top">Notes <%if (NotesServicesArray[i][8] != null) Response.Write("("+NotesServicesArray[i][8]+")")%>&nbsp;</td>
		<td valign="top" nowrap><%=(NotesServicesArray[i][6])%>&nbsp;</td>
<%		
if (Request.QueryString("ShowFS")=="1") {
%>				
		<td valign="top" nowrap>&nbsp;</td>
<%
}
%>
		<td valign="top" nowrap><%=(NotesServicesArray[i][10])%>&nbsp;</td>
    </tr>
<%
	} else {
%>
    <tr>
		<td valign="top"><a href="m012e0501.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intSrv_Note_id=<%=NotesServicesArray[i][1]%>&intService_Req_id=<%=NotesServicesArray[i][3]%>"><%=FilterDate(NotesServicesArray[i][0])%></a>&nbsp;</td>
		<td valign="top">Service Request <%if (NotesServicesArray[i][7] != null) Response.Write("("+NotesServicesArray[i][7]+")")%>&nbsp;</td>
		<td valign="top" nowrap><%=(NotesServicesArray[i][6])%>&nbsp;</td>
<%		
if (Request.QueryString("ShowFS")=="1") {
%>				
		<td valign="top" nowrap><%=(NotesServicesArray[i][9])%>&nbsp;</td>
<%
}
%>
		<td valign="top" nowrap><%=(NotesServicesArray[i][10])%>&nbsp;</td>
    </tr>
<%
	}
}
%>
</table>
</body>
</html>
<%
rsNotesServices.Close();
%>
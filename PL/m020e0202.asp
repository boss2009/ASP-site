<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsResponseHistory = Server.CreateObject("ADODB.Recordset");
rsResponseHistory.ActiveConnection = MM_cnnASP02_STRING;
rsResponseHistory.Source = "{call dbo.cp_pjt_responses("+Request.QueryString("intIssue_id")+",'',0,0,0,0,'',0,0,0,'',0,'',0,'Q',0)}";
rsResponseHistory.CursorType = 0;
rsResponseHistory.CursorLocation = 2;
rsResponseHistory.LockType = 3;
rsResponseHistory.Open();
%>
<html>
<head>
	<title>Response History</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
	  	 	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
</head>
<body>
<h5>Response History</h5>
<% 
while (!rsResponseHistory.EOF){
%>
<hr>
<table cellpadding="2" cellspacing="1">
    <tr>       
		<td class="headrow" align="left" nowrap>Created By:</td>      
		<td style="border: solid 1px #cccccc" width="231"><%=(rsResponseHistory.Fields.Item("chvSubmitted_by").Value)%>&nbsp;</td>      
		<td class="headrow" align="left" width="105">Assigned To:</td>
		<td style="border: solid 1px #cccccc" width="217"><%=(rsResponseHistory.Fields.Item("chvAssigned_to").Value)%>&nbsp;</td>	  
    </tr>
    <tr>       
		<td class="headrow" align="left" nowrap>Date:</td>      
		<td style="border: solid 1px #cccccc"><%=(rsResponseHistory.Fields.Item("dtsDate_response").Value)%>&nbsp;</td>      
		<td class="headrow" align="left">Version:</td>      
		<td style="border: solid 1px #cccccc"><%=(rsResponseHistory.Fields.Item("ncvVersion").Value)%>&nbsp;</td>
    </tr>
    <tr>       
		<td class="headrow" align="left" nowrap>Tested:</td>      
		<td style="border: solid 1px #cccccc"><%=(rsResponseHistory.Fields.Item("bitTested").Value)%>&nbsp;</td>      
		<td class="headrow" align="left">Approved:</td>      
		<td style="border: solid 1px #cccccc"><%=(rsResponseHistory.Fields.Item("bitApproved").Value)%>&nbsp;</td>
    </tr>
    <tr>       
		<td class="headrow" align="left" nowrap>Priority:</td>      
		<td style="border: solid 1px #cccccc"><%=(rsResponseHistory.Fields.Item("chvPriority").Value)%>&nbsp;</td>      
		<td class="headrow" align="left">Status:</td>      
		<td style="border: solid 1px #cccccc"><%=(rsResponseHistory.Fields.Item("chvStatus").Value)%>&nbsp;</td>
    </tr>
    <tr>       
		<td class="headrow" valign="top" align="left">Response:</td>
		<td colspan="3"><textarea cols="75" rows="7" readonly tabindex="1" accesskey="F"><%=(rsResponseHistory.Fields.Item("chvDescription").Value)%></textarea></td>
    </tr>
</table>
<%
	rsResponseHistory.MoveNext();
}
%>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Close" onClick="window.close();" tabindex="2" class="btnstyle"></td>
	</tr>
</table>
</body>
</html>
<%
rsResponseHistory.Close();
%>
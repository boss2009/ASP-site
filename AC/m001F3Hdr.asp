<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_FrmHdr(1,"+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

var rsClient2 = Server.CreateObject("ADODB.Recordset");
rsClient2.ActiveConnection = MM_cnnASP02_STRING;
rsClient2.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient2.CursorType = 0;
rsClient2.CursorLocation = 2;
rsClient2.LockType = 3;
rsClient2.Open();
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
	<title>Client Header Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 570px; top: 10px"> 
  <table cellspacing="1" cellpadding="1">
    <tr> 
      <td><b>Client:</b></td>
      <td width=200><%=(rsClient.Fields.Item("Name").Value)%></td>
      <td><b>SIN:</b></td>
      <td><%=FormatSIN(Trim(rsClient2.Fields.Item("chrSIN_no").Value))%></td>
    </tr>
    <tr> 
      <td><b>Disability:</b></td>
      <td><%=(rsClient.Fields.Item("Disability").Value)%></td>
      <td><b>Case Manager:</b></td>
      <td><%=(rsClient.Fields.Item("Case Manager").Value)%></td>
    </tr>
    <tr> 
      <td><b>Status:</b></td>
      <td><%=(rsClient.Fields.Item("Status").Value)%></td>
      <td><b>First Referral Date:</b></td>
      <td><%=FilterDate(rsClient.Fields.Item("Referral_Date").Value)%></td>
    </tr>
    <tr> 
      <td><b>Region:</b></td>
      <td><%=(rsClient.Fields.Item("Region").Value)%></td>
      <td><b>Most Recent Referral:</b></td>
      <td><%=FilterDate(rsClient.Fields.Item("Re-referral_Date").Value)%></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsClient.Close();
%>
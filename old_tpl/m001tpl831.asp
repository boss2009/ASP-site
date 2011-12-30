<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/msword" %>

<%
var rsClient__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ rsClient__intpAdult_id.replace(/'/g, "''") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
%>
<html>
<head>
<title>User's Declaration</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aspform.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="480" border="0">
  <tr> 
    <td colspan="3" class="HdrReg"> 
      <div align="center">User's Declaration</div>
    </td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411">&nbsp;</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3" class="HdrReg">I, <%=(rsClient.Fields.Item("chvName").Value)%>, understand that should technical problems arise with my computer, 
      I will inform my referring agent and contact Assistive Technology - British Columbia first 
      for troubleshooting assistance. I recognize that I may be charged for costs 
      incurred in resolving problems resulting from:</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411">&nbsp;</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411" class="HdrReg">unauthorized User</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411" class="HdrReg">unauthorized software installation</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411" class="HdrReg">unauthorized customization/reconfiguration 
      of the computer's operating system</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411" class="HdrReg">unauthorized customization/reconfiguration 
      of the adaptive software</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411" class="HdrReg">unauthorized reformatting of the HD</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411" class="HdrReg">unauthorized installation of hardware</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411" class="HdrReg">not keeping original shipping boxes and packing 
      so it can be reused should equipment need to be returned</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411">&nbsp;</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411">&nbsp;</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2" class="HdrReg">.Signature:</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411">&nbsp;</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411">&nbsp;</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2" class="HdrReg">Witness: </td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411">&nbsp;</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td width="411">&nbsp;</td>
    <td width="19">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2" class="HdrReg">Date:</td>
    <td width="19">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
rsClient.Close();
%>

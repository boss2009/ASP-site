<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/msword" %>
<%
var rsContact__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsContact__intContact_id = String(Request.QueryString("intContact_id"));
var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_Cntct_Detail("+ rsContact__intpAdult_id.replace(/'/g, "''") + ","+ rsContact__intContact_id.replace(/'/g, "''") + ")}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

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
<title>Enrollment - Cont.</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aspform.css" type="text/css">
</head>
<body text="#000000" bgcolor="#FFFFFF">
<table width="725" border="0" class="HdrReg">
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td nowrap colspan="4"><%=(rsContact.Fields.Item("chvTitle").Value)%> <%=(rsContact.Fields.Item("chvFst_Name").Value)%> <%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvJob_title").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvWork_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvAddress").Value)%></td>
  </tr>
  <tr> 
    <td nowrap colspan="4"><%=(rsContact.Fields.Item("chvCity").Value)%> <%=(rsContact.Fields.Item("chvProv").Value)%> <%=(rsContact.Fields.Item("chvCntry_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvPostal_Zip").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Dear <%=(rsContact.Fields.Item("chvTitle").Value)%> <%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td width="16">re: </td>
    <td width="476">Enrollment Verification</td>
    <td width="185">&nbsp;</td>
    <td width="30">&nbsp;</td>
  </tr>
  <tr> 
    <td width="16">&nbsp;</td>
    <td width="476"><%=(rsClient.Fields.Item("chvName").Value)%></td>
    <td width="185">&nbsp;</td>
    <td width="30">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"> 
      <div align="left">As outlined in the Conditions of Equipment Loan, we require 
        current confirmation of continued enrollment for each semester. At this 
        time we are asking if the individual(s) listed below is presently enrolled, 
        in a minimum of two credit courses, as equipment is on loan for educational 
        purposes only. In order to receive continued services, we would like this 
        information as soon as possible</div>
    </td>
    <td width="30">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">Your attention to this matter is greatly appreciated.</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Yours truly,</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">D T Chan ,</td>
  </tr>
  <tr> 
    <td colspan="4">CEO</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td width="16">cc: </td>
    <td colspan="3">VRS Consultant</td>
  </tr>
  <tr> 
    <td width="16">&nbsp;</td>
    <td colspan="3"> Senior Coordinator, Ministry Human Resources</td>
  </tr>
  <tr> 
    <td width="16">&nbsp;</td>
    <td colspan="3"> Special Programs, AVED, Student 
      Services Branch</td>
  </tr>
  <tr> 
    <td width="16">&nbsp;</td>
    <td colspan="3">Regional Coordinator</td>
  </tr>
</table>
</body>
</html>
<%
rsContact.Close();
%>
<%
rsClient.Close();
%>

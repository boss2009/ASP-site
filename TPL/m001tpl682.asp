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

var rsLoadReq__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsLoadReq = Server.CreateObject("ADODB.Recordset");
rsLoadReq.ActiveConnection = MM_cnnASP02_STRING;
rsLoadReq.Source = "{call dbo.cp_Loan_Request("+ rsLoadReq__intpAdult_id.replace(/'/g, "''") + ")}";
rsLoadReq.CursorType = 0;
rsLoadReq.CursorLocation = 2;
rsLoadReq.LockType = 3;
rsLoadReq.Open();
%>
<html>
<head>
<title>Default - P/S Equipment Loan</title>
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
    <td width="42">re: </td>
    <td colspan="3">DEFAULT STATUS</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td colspan="3"><%=(rsClient.Fields.Item("chvName").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">      <div align="left"><%=(rsClient.Fields.Item("chvName").Value)%> is in default of the established equipment loan plan dated 
        [--date--] from the Sirius Innovations Inc.. We have made numerous requests 
        and attempts to retrieve the equipment. I am requesting that this letter 
        be placed in both the VRS file and the [POST-SECONDARY] file for <%=(rsClient.Fields.Item("chvName").Value)%>, as a reminder that <%=(rsClient.Fields.Item("chvFst_Name").Value)%> is not in good standing with Sirius Innovations Inc.</div>
    </td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">This equipment in on loan for educational purposes only and 
      <%=(rsClient.Fields.Item("chvFst_Name").Value)%> is no longer enrolled and attending a post secondary institution, 
      therefore,<%=(rsClient.Fields.Item("chvFst_Name").Value)%> no longer meets eligibility criteria. Sirius Innovations Inc. 
      will add a note in our database that the equipment listed below is out of 
      inventory and that the equipment is in default status. We will close our 
      file at this point</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <% while (!rsLoadReq.EOF) { %>
  <tr> 
    <td width="42">&nbsp; 
    </td>
    <td width="59">&nbsp;</td>
    <td width="557"><%=(rsLoadReq.Fields.Item("chvInventory_Name").Value)%></td>
    <td width="49">&nbsp;</td>
  </tr>
  <%
  rsLoadReq.MoveNext();
}
%>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">We require either arrangements to be made for the return of 
      equipment, or payment in the amount of dollars ($10.00) to rescind the provincial 
      default status. Please call me if you wish to discuss this further.</td>
    <td width="49">&nbsp;</td>
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
    <td width="42">cc: </td>
    <td colspan="3"><%=(rsClient.Fields.Item("chvName").Value)%></td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td colspan="3">Senior Coordinator, Ministry Human Resources</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td colspan="3">A.Coordinator, Special Programs, AVED, Student 
      Services Branch</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
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
<%
rsLoadReq.Close();
%>


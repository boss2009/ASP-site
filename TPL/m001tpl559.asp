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
%>
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
<title>Accept - CSG Plan</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aspform.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="570" border="0" class="HdrReg">
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6"><%=(rsContact.Fields.Item("chvTitle").Value)%> <%=(rsContact.Fields.Item("chvFullName").Value)%></td>
  </tr>
  <tr> 
    <td colspan="6"><%=(rsContact.Fields.Item("chvJob_title").Value)%></td>
  </tr>
  <tr> 
    <td colspan="6"><%=(rsContact.Fields.Item("chvWork_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="6"><%=(rsContact.Fields.Item("chvAddress").Value)%></td>
  </tr>
  <tr> 
    <td colspan="6"><%=(rsContact.Fields.Item("chvCity").Value)%>, <%=(rsContact.Fields.Item("chvProv").Value)%>, <%=(rsContact.Fields.Item("chvCntry_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="6"><%=(rsContact.Fields.Item("chvPostal_Zip").Value)%></td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">Dear <%=(rsContact.Fields.Item("chvTitle").Value)%> <%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">Re:</td>
    <td colspan="5"><%=(rsClient.Fields.Item("chvName").Value)%></td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">As Manager of the Sirius Innovations Inc., I am writing you 
      with regards to <%=(rsClient.Fields.Item("chvName").Value)%> 's referral application received by this office. We are pleased 
      to inform you that <%=(rsClient.Fields.Item("chvFst_Name").Value)%> has been accepted for services through the Sirius Innovations Inc.</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">The plan we are recommending is as follows:</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td colspan="4">Using funds eligible through Canada Study Grant Program to 
      purchase:</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25" height="17">&nbsp;</td>
    <td colspan="4" height="17">Using funds eligible through APSD Service Grant 
      to complete the purchase of equipment.<br>
    </td>
    <td width="22" height="17">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">If the student agrees, please have the student sign the attached 
      waiver form (see attached) and return to Sirius Innovations Inc for implementation 
      of this plan. Please inform <%=(rsClient.Fields.Item("chvFst_Name").Value)%> that this plan is subject to federal income tax regulations.</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">This funding is to allow <%=(rsClient.Fields.Item("chvFst_Name").Value)%> to purchase equipment during current study period for the CSG 
      program year August 01, 2004 and ending July 31, 2005</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">We will then cancel the current loan agreement once funding 
      has been received</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">The equipment is now on order and will be shipped as soon 
      as possible when we receive the following:</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td colspan="4"> verification of enrollment</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td colspan="4">signed Conditions of Equipment Ownership/Loan form (see attached)</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">Si2 personnel will then contact you to make arrangements for 
      delivery of equipment and any necessary training.</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">If you have any further questions or concerns regarding this 
      decision, please feel free to contact me at (604) 959-8188 </td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">Thank you for this referral and we hope that the Sirius Innovations Inc. can be of assistance to you in the future</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2">Yours truly,</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">D T Chan </td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">CEO</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td width="193">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="102">&nbsp;</td>
    <td width="101">&nbsp;</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">cc</td>
    <td colspan="4"><%=(rsClient.Fields.Item("chvName").Value)%></td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="25">&nbsp;</td>
    <td colspan="4">  A Coordinator, Special Programs, AVED, 
      Student Services</td>
    <td width="22">&nbsp;</td>
  </tr>
  <tr>
    <td width="25">&nbsp;</td>
    <td colspan="4">Regional Coordinator</td>
    <td width="22">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
rsContact.Close();
rsClient.Close();
%>

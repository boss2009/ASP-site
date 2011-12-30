<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/msword" %>

<%
var rsContact__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsContact__intContact_id = String(Request.QueryString("intContact_id"));
var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_Cntct_Detail_2("+ rsContact__intpAdult_id.replace(/'/g, "''") + ","+ rsContact__intContact_id.replace(/'/g, "''") + ")}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

var rsClient__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client_Detail("+ rsClient__intpAdult_id.replace(/'/g, "''") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
var rsClient_numRows = 0;

var rsASP = Server.CreateObject("ADODB.Recordset");
rsASP.ActiveConnection = MM_cnnASP02_STRING;
rsASP.Source = "{call dbo.cp_Idv_CmpyInfo(777)}";
rsASP.CursorType = 0;
rsASP.CursorLocation = 2;
rsASP.LockType = 3;
rsASP.Open();
%>
<html>
<head>
<title>Follow Up - Form Education</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aspform.css" type="text/css">
</head>
<body text="#000000" bgcolor="#FFFFFF">
<table width="630" border="0" class="HdrReg">
  <tr> 
    <td colspan="4"> 
      <table width="624" border="0">
        <tr> 
          <td width="71" rowspan="3"> 
            <div align="left"><img src="../i/CA.gif" width="68" height="50"></div>
          </td>
          <td width="543" class="aspxsmallCap"> 
            <div align="center"><%=(rsASP.Fields.Item("chvName").Value)%></div>
          </td>
        </tr>
        <tr> 
          <td width="543" class="aspxsmallCap"><%=(rsASP.Fields.Item("chvAddress").Value)%> , <%=(rsASP.Fields.Item("chvCity").Value)%> , <%=(rsASP.Fields.Item("chvProvince").Value)%> , <%=(rsASP.Fields.Item("chvPostal").Value)%></td>
        </tr>
        <tr> 
          <td width="543"> 
            <table width="541" border="0">
              <tr> 
                <td width="30" class="aspxsmallCap">Tel:</td>
                <td width="233" nowrap class="aspxsmallCap"><%=(rsASP.Fields.Item("chvHm_Areacd").Value)%>- <%=(rsASP.Fields.Item("chvHm_no").Value)%></td>
                <td width="26" class="aspxsmallCap"> Fax:</td>
                <td width="234" class="aspxsmallCap"><%=(rsASP.Fields.Item("chvFax_Areacd").Value)%>- <%=(rsASP.Fields.Item("chvFax_no").Value)%></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td width="71" rowspan="4">&nbsp;</td>
          <td width="543">&nbsp;</td>
        </tr>
        <tr> 
          <td width="543" class="aspSB"> 
            <div align="center">Follow-Up Form: Educational</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4" class="aspSB">1. APPLICANT INFORMATION</td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsClient.Fields.Item("chvName").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsClient.Fields.Item("chvAddress").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsClient.Fields.Item("chvCity").Value)%> , <%=(rsClient.Fields.Item("chrprvst_abbv").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsClient.Fields.Item("chvPostal_zip").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4" nowrap><%=(rsClient.Fields.Item("chvPhoneName").Value)%> : <%=(rsClient.Fields.Item("chvPhone1_Arcd").Value)%> - <%=(rsClient.Fields.Item("chvPhone1_Num").Value)%> - <%=(rsClient.Fields.Item("chvPhone1_Ext").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvFullName").Value)%></td>
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
    <td colspan="4"><%=(rsContact.Fields.Item("chvCity").Value)%> , <%=(rsContact.Fields.Item("chvProv").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvPostal_Zip").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvPhoneType_1").Value)%> : <%=(rsContact.Fields.Item("chvPhone1_Arcd").Value)%> - <%=(rsContact.Fields.Item("chvPhone1_Num").Value)%> <%=(rsContact.Fields.Item("chvPhone1_Ext").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvPhoneType_2").Value)%> : <%=(rsContact.Fields.Item("chvPhone2_Arcd").Value)%> - <%=(rsContact.Fields.Item("chvPhone2_Num").Value)%> <%=(rsContact.Fields.Item("chvPhone2_Ext").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">2. QUESTIONS</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Have the services and/or equipment provided met the goals 
      identified in the educational plan? If no, please explain</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="aspSB">Yes</td>
    <td class="aspSB">No</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Has the educational goal for the student changed?</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="aspSB">Yes</td>
    <td class="aspSB">No</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Has the nature of the disability changed?</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="aspSB">Yes</td>
    <td class="aspSB">No</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">What is the expected completion dated of the current educational 
      program?</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">We verify that the student is successfully completing courses 
      and will be enrolled in September, 2004. Therefore, we request the continued 
      loan of the equipment through </td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312">&nbsp;</td>
    <td width="322" nowrap class="aspSB">FALL 2000</td>
    <td width="189" class="aspSB">SPRING 2001</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312">&nbsp;</td>
    <td width="322">&nbsp;</td>
    <td width="189">&nbsp;</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312">&nbsp;</td>
    <td width="322">&nbsp;</td>
    <td width="189">&nbsp;</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312" class="aspSB">Signature Referring Agent</td>
    <td width="322">&nbsp;</td>
    <td width="189" class="aspSB">Date</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312">&nbsp;</td>
    <td width="322">&nbsp;</td>
    <td width="189">&nbsp;</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312">&nbsp;</td>
    <td width="322">&nbsp;</td>
    <td width="189">&nbsp;</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312">&nbsp;</td>
    <td width="322">&nbsp;</td>
    <td width="189">&nbsp;</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312" class="aspSB">Signature Client</td>
    <td width="322">&nbsp;</td>
    <td width="189" class="aspSB">Date</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312">&nbsp;</td>
    <td width="322">&nbsp;</td>
    <td width="189">&nbsp;</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td width="312">&nbsp;</td>
    <td width="322">&nbsp;</td>
    <td width="189">&nbsp;</td>
    <td width="49">&nbsp;</td>
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
rsASP.Close();
%>

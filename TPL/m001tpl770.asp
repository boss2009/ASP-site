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

var rsASP = Server.CreateObject("ADODB.Recordset");
rsASP.ActiveConnection = MM_cnnASP02_STRING;
rsASP.Source = "{call dbo.cp_Idv_CmpyInfo(777)}";
rsASP.CursorType = 0;
rsASP.CursorLocation = 2;
rsASP.LockType = 3;
rsASP.Open();

var rsLoanReq__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsLoanReq = Server.CreateObject("ADODB.Recordset");
rsLoanReq.ActiveConnection = MM_cnnASP02_STRING;
rsLoanReq.Source = "{call dbo.cp_Loan_Request("+ rsLoanReq__intpAdult_id.replace(/'/g, "''") + ")}";
rsLoanReq.CursorType = 0;
rsLoanReq.CursorLocation = 2;
rsLoanReq.LockType = 3;
rsLoanReq.Open();
var rsLoanReq_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsLoanReq_numRows += Repeat1__numRows;
%>
<html>
<head>
<title>Follow Up - Form Vocational</title>
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
            <div align="left"><img src="../i/asplogo.jpg" width="70" height="70"></div>
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
                <td width="30" class="aspxsmallCap" height="22">Tel:</td>
                <td width="233" nowrap class="aspxsmallCap" height="22"><%=(rsASP.Fields.Item("chvHm_Areacd").Value)%>- <%=(rsASP.Fields.Item("chvHm_no").Value)%></td>
                <td width="26" class="aspxsmallCap" height="22"> Fax:</td>
                <td width="234" class="aspxsmallCap" height="22"><%=(rsASP.Fields.Item("chvFax_Areacd").Value)%>- <%=(rsASP.Fields.Item("chvFax_no").Value)%></td>
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
            <div align="center">Follow-Up Form: Vocational</div>
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
    <td colspan="4" class="aspxsmallCap">2.EQUIPMENT INFORMATION</td>
  </tr>
  <tr> 
    <td colspan="3" nowrap>&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <% while ((Repeat1__numRows-- != 0) && (!rsLoanReq.EOF)) { %>
  <tr> 
    <td colspan="4"> 
      <table width="650" border="0">
        <tr> 
          <td width="61">&nbsp;</td>
          <td width="579" class="HdrReg"><%=(rsLoanReq.Fields.Item("chvInventory_Name").Value)%></td>
        </tr>
      </table>
    </td>
  </tr>
  <%
  Repeat1__index++;
  rsLoanReq.MoveNext();
}
%>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4" class="aspxsmallCap">3. QUESTIONS</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Have the services and/or equipment provided met the goals 
      identified in the vocational plan? If no, please explain.</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="aspSB" width="344">Yes</td>
    <td class="aspSB" width="199">No</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Has the vocational goal changed?</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="aspSB" width="344">Yes</td>
    <td class="aspSB" width="199">No</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Has the nature of the disability changed?</td>
  </tr>
  <tr> 
    <td height="20">&nbsp;</td>
    <td class="aspSB" width="344" height="20">Yes</td>
    <td class="aspSB" width="199" height="20">No</td>
    <td width="42" height="20">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">Adult Service Program will support the loaned equipment for 
      approximately one year to ensure the equipment has resolved the barriers 
      to employment caused by the disability. The present one year loan period 
      has expires on:</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">Is there a plan in place for the buyout? </td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344" nowrap class="aspSB">Yes</td>
    <td width="199" class="aspSB">No</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344" nowrap class="aspSB">&nbsp;</td>
    <td width="199" class="aspSB">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344" nowrap class="aspSB">&nbsp;</td>
    <td width="199" class="aspSB">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td nowrap class="aspSB">Signature Referring Agent</td>
    <td width="344">&nbsp;</td>
    <td width="199" class="aspSB">Date</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td nowrap class="aspSB">Signature Client</td>
    <td width="344">&nbsp;</td>
    <td width="199" class="aspSB">Date</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="344">&nbsp;</td>
    <td width="199">&nbsp;</td>
    <td width="42">&nbsp;</td>
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
<%
rsLoanReq.Close();
%>


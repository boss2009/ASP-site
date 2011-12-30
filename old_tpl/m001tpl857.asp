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

var rsClient__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ rsClient__intpAdult_id.replace(/'/g, "''") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
var rsClient_numRows = 0;
%>
<html>
<head>
<title>PILAT Loan Review</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aspform.css" type="text/css">
</head>
<body text="#000000" bgcolor="#FFFFFF">
<table width="630" border="0" class="HdrReg">
  <tr> 
    <td colspan="4"> 
      <table width="652" border="0">
        <tr> 
          <td rowspan="3" colspan="4"> 
            <div align="left"><img src="../i/asplogo2.jpg" width="73" height="70"></div>
          </td>
          <td width="571" class="aspxsmallCap"> 
            <div align="center"><%=(rsASP.Fields.Item("chvName").Value)%></div>
          </td>
        </tr>
        <tr> 
          <td width="571" class="aspxsmallCap"><%=(rsASP.Fields.Item("chvAddress").Value)%> , <%=(rsASP.Fields.Item("chvCity").Value)%> , <%=(rsASP.Fields.Item("chvProvince").Value)%> , <%=(rsASP.Fields.Item("chvPostal").Value)%></td>
        </tr>
        <tr> 
          <td width="571"> 
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
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4" class="aspSB"><%=(rsContact.Fields.Item("chvTitle").Value)%><%=(rsContact.Fields.Item("chvFst_Name").Value)%> <%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
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
    <td colspan="4"><%=(rsContact.Fields.Item("chvCity").Value)%> <%=(rsContact.Fields.Item("chvProv").Value)%> <%=(rsContact.Fields.Item("chvCntry_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="4" nowrap><%=(rsContact.Fields.Item("chvPostal_Zip").Value)%></td>
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
    <td colspan="4">Re: REVIEW OF PILAT LOAN AGREEMENT</td>
  </tr>
  <tr> 
    <td colspan="4">
      <table width="657" border="0">
        <tr>
          <td width="614" class="HdrReg"><%=(rsClient.Fields.Item("chvName").Value)%></td>
          <td width="32">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4"> 
      <table width="655" border="0">
        <tr> 
          <td width="613" class="HdrReg">Referring Agent Organization has a Low 
            Utilization institutional loan of</td>
          <td width="31">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <% while (!rsLoanReq.EOF) { %>
  <tr> 
    <td colspan="4"> 
      <table width="656" border="0">
        <tr> 
          <td width="47" height="17">&nbsp;</td>
          <td class="HdrReg" height="17"><%=(rsLoanReq.Fields.Item("chvInventory_Name").Value)%></td>
          <td width="31" height="17">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <%
  rsLoanReq.MoveNext();
}
%>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4"> 
      <table width="653" border="0">
        <tr> 
          <td width="614" class="HdrReg"> from the Assistive Technology - British Columbia through 
            PILAT. The PILAT Policy indicates that the loan agreement be reviewed 
            annually to ensure that the loan criteria and the student&Otilde;s 
            educational goals are being met. </td>
          <td width="28" class="HdrReg">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4" class="HdrReg"> 
      <table width="654" border="0">
        <tr> 
          <td width="617" class="HdrReg">Please contact the office to make arrangements 
            for a Review, by contacting your Adult Services case manager to set 
            a date for the Review. If you have questions or concerns regarding 
            the loan or the Review, please feel free to contact me at (604) 264-8295</td>
          <td width="26">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">Thank you for your attention to this matter</td>
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
    <td colspan="4">Yours truly,</td>
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
    <td colspan="4">Gladys Loewen,</td>
  </tr>
  <tr> 
    <td colspan="4">Program Manager</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
rsContact.Close();
rsASP.Close();
rsLoanReq.Close();
rsClient.Close();
%>

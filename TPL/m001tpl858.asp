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

var rsLoadReq__intpAdult_id = String(Request.QueryString("intAdult_id"));
var rsLoadReq = Server.CreateObject("ADODB.Recordset");
rsLoadReq.ActiveConnection = MM_cnnASP02_STRING;
rsLoadReq.Source = "{call dbo.cp_Loan_Request("+ rsLoadReq__intpAdult_id.replace(/'/g, "''") + ")}";
rsLoadReq.CursorType = 0;
rsLoadReq.CursorLocation = 2;
rsLoadReq.LockType = 3;
rsLoadReq.Open();
var rsLoadReq_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsLoadReq_numRows += Repeat1__numRows;
%>

<html>
<head>
<title>PILAT Accept</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aspform.css" type="text/css">
</head>
<body text="#000000" bgcolor="#FFFFFF">
<table width="725" border="0" class="HdrReg">
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4"><%=(rsContact.Fields.Item("chvTitle").Value)%> <%=(rsContact.Fields.Item("chvFst_Name").Value)%> <%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
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
    <td colspan="4"><%=(rsContact.Fields.Item("chvCity").Value)%> <%=(rsContact.Fields.Item("chvProv").Value)%></td>
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
    <td colspan="3">: PROGRAM FOR INSTITUTIONAL LOANS OF Sirius Innovations Inc.</td>
  </tr>
  <tr> 
    <td width="42">&nbsp;</td>
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"> 
      <div align="left">As Manager of the Sirius Innovations Inc. I am writing 
        you with regards to the (post-secondary institution) application for a 
        Temp loan. We are pleased to inform you that (post-seconary Institution) 
        has been accepted for an institutional loan through the Sirius Innovations Inc. The following is a list of the hardware/software to be loaned:</div>
    </td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <% while ((Repeat1__numRows-- != 0) && (!rsLoadReq.EOF)) { %>
  <tr> 
    <td width="42">&nbsp; 
    </td>
    <td width="59">&nbsp;</td>
    <td width="557"><%=(rsLoadReq.Fields.Item("chvInventory_Name").Value)%></td>
    <td width="49">&nbsp;</td>
  </tr>
  <%
  Repeat1__index++;
  rsLoadReq.MoveNext();
}
%>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">The loan request is now being processed and will be shipped 
      as soon as possible when we receive the following:</td>
    <td width="49">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>( ) Temp Conditions of Loan form</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>( ) Student Information as per Application</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">Also note that the expected return date of the loaned hardware/software 
      is __________________________. Please contact our office before the due 
      date to arrange for the return of the hardware/software or to discuss any 
      extenuating circumstances requiring an extension of the loan. </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">If you have any questions or concerns regarding this decision, 
      please feel free to contact me a (604) 959-8188 </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">Thank you for this referral and we hope that the Sirius Innovations Inc. can be of assistance to you in the future</td>
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
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4">CEO</td>
  </tr>
  <tr> 
    <td width="42">cc: </td>
    <td colspan="3"> MAE, Post Secondary Education Division</td>
  </tr>
</table>
</body>
</html>
<%
rsContact.Close();
rsLoadReq.Close();
%>

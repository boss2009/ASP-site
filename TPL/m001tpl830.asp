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
var rsLoadReq_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsLoadReq_numRows += Repeat1__numRows;
%>
<html>
<head>
<title>VRS Buyout Offer Revised</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aspform.css" type="text/css">
</head>
<body text="#000000" bgcolor="#FFFFFF">
<table width="725" border="0" class="HdrReg">
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td nowrap colspan="6"><%=(rsContact.Fields.Item("chvTitle").Value)%> <%=(rsContact.Fields.Item("chvFst_Name").Value)%> <%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
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
    <td nowrap colspan="6"><%=(rsContact.Fields.Item("chvCity").Value)%> <%=(rsContact.Fields.Item("chvProv").Value)%> <%=(rsContact.Fields.Item("chvCntry_Name").Value)%></td>
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
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td width="41">re: </td>
    <td width="8">&nbsp;</td>
    <td colspan="4">EQUIPMENT PURCHASE PLAN</td>
  </tr>
  <tr> 
    <td width="41">&nbsp;</td>
    <td width="8">&nbsp;</td>
    <td colspan="4"><%=(rsClient.Fields.Item("chvName").Value)%></td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5"> 
      <div align="left">Further to our letter dated [date] outlining the equipment 
        purchase plan for the equipment on loan, in accordance with the MSDES/ASP 
        policy regarding the support for loaned equipment for approximately o 
        n year on employment sites, to date this equipment purchase plan has not 
        been initiated. At this time we require that either the equipment purchase 
        plan be initiated or arrangements made for the return of this equipment. 
        The following lists the revised offer we are able to make:</div>
    </td>
    <td width="122">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <% while ((Repeat1__numRows-- != 0) && (!rsLoadReq.EOF)) { %>
  <tr> 
    <td width="41">&nbsp; 
    </td>
    <td width="8">&nbsp;</td>
    <td colspan="4"><%=(rsLoadReq.Fields.Item("chvInventory_Name").Value)%></td>
  </tr>
  <%
  Repeat1__index++;
  rsLoadReq.MoveNext();
}
%>
  <tr> 
    <td width="41">&nbsp;</td>
    <td width="8">&nbsp;</td>
    <td width="75">&nbsp;</td>
    <td width="127">Subtotal</td>
    <td width="326">&nbsp;</td>
    <td width="122">&nbsp;</td>
  </tr>
  <tr> 
    <td width="41">&nbsp;</td>
    <td width="8">&nbsp;</td>
    <td width="75">&nbsp;</td>
    <td width="127">LESS: discount</td>
    <td width="326">&nbsp;</td>
    <td width="122">&nbsp;</td>
  </tr>
  <tr> 
    <td width="41">&nbsp;</td>
    <td width="8">&nbsp;</td>
    <td width="75">&nbsp;</td>
    <td width="127">Buyout Offer</td>
    <td width="326">&nbsp;</td>
    <td width="122">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="5">If you agree to the purchase price, please contact the office 
      to make arrangements to implement this equipment purchase plan, at your 
      earliest convenience. We will forward an invoice for the purchase of the 
      equipment upon request. We will cancel the current loan agreement once payment 
      has been received.</td>
    <td width="122">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">Thank you for your attention to this matter. If you have any 
      further questions or concerns, do not hesitate to call.</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">Yours truly,</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">D T Chan ,</td>
  </tr>
  <tr> 
    <td colspan="6">CEO</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="6">cc: <%=(rsClient.Fields.Item("chvName").Value)%></td>
  </tr>
  <tr> 
    <td colspan="6">&nbsp;</td>
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


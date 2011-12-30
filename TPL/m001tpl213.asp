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
<title>VRS Loan Extension</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../css/aspform.css" type="text/css">
</head>
<body text="#000000" bgcolor="#FFFFFF">
<table width="725" border="0" class="HdrReg">
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td nowrap colspan="3"><%=(rsContact.Fields.Item("chvTitle").Value)%> <%=(rsContact.Fields.Item("chvFst_Name").Value)%> <%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="3"><%=(rsContact.Fields.Item("chvJob_title").Value)%></td>
  </tr>
  <tr> 
    <td colspan="3"><%=(rsContact.Fields.Item("chvWork_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="3"><%=(rsContact.Fields.Item("chvAddress").Value)%></td>
  </tr>
  <tr> 
    <td nowrap colspan="3"><%=(rsContact.Fields.Item("chvCity").Value)%> <%=(rsContact.Fields.Item("chvProv").Value)%> <%=(rsContact.Fields.Item("chvCntry_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="3"><%=(rsContact.Fields.Item("chvPostal_Zip").Value)%></td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">Dear <%=(rsContact.Fields.Item("chvTitle").Value)%> <%=(rsContact.Fields.Item("chvLst_Name").Value)%></td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"> 
      <table width="480" border="0">
        <tr> 
          <td width="14" class="HdrReg">re:</td>
          <td width="456" class="HdrReg"><%=(rsClient.Fields.Item("chvName").Value)%></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="3"> 
      <table width="725" border="0">
        <tr> 
          <td width="701" class="HdrReg">As Manager of the Sirius Innovations Inc., 
            I am writing you with regards to your request to extend <%=(rsClient.Fields.Item("chvName").Value)%>'s equipment loan. We are pleased to inform you that Si2 is able to extend the loan for:</td>
          <td width="14">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3"> 
      <table width="728" border="0">
        <tr> 
          <td width="30">&nbsp;</td>
          <td width="688" class="HdrReg">xxx months to allow for <%=(rsClient.Fields.Item("chvName").Value)%>'s health to stabilize</td>
        </tr>
        <tr> 
          <td width="30">&nbsp;</td>
          <td width="688" class="HdrReg">xxx months to secure other employment 
            or PSTP position</td>
        </tr>
        <tr> 
          <td width="30">&nbsp;</td>
          <td width="688" class="HdrReg">xxx months</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3" height="30">
      <table width="709" border="0">
        <tr> 
          <td width="703" class="HdrReg">After this period, we will require written 
            confirmation that <%=(rsClient.Fields.Item("chvName").Value)%> meets Si2's eligibility criteria. A new CIP may also 
            be required should <%=(rsClient.Fields.Item("chvName").Value)%> find 
            other employment. Otherwise, arrangements will need to be made to 
            return the equipment.</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">If you have any further questions or concerns regarding this 
      decision, please feel free to contact me at (604)-959-8188 </td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">Yours truly,</td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">D T Chan ,</td>
  </tr>
  <tr> 
    <td colspan="3">CEO</td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3">cc: <%=(rsClient.Fields.Item("chvName").Value)%></td>
  </tr>
  <tr> 
    <td colspan="3">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
rsContact.Close();
rsClient.Close();
%>

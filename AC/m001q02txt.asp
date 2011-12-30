<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "text/plain" %>
<%
var rsClient__inspSrtBy = String(Request.QueryString("inspSrtBy"));
var rsClient__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
var rsClient__chvFilter = "dtsRefral_date >= '01/01/1900'";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsClient__chvFilter = String(Request.QueryString("chvFilter"));
}
var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Adult_Client2E("+ rsClient__inspSrtBy.replace(/'/g, "''") + ","+ rsClient__inspSrtOrd.replace(/'/g, "''") + ",'"+ rsClient__chvFilter.replace(/'/g, "''") + "')}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
%>
"Last Name","First Name","ASP ID","SIN","Primary Disability","Referral Date","Re-referral Date","Referral Type","Referring Agent","Status","Case Manager"<% while (!rsClient.EOF) { %>
"<%=(rsClient.Fields.Item("chvLst_Name").Value)%>","<%=(rsClient.Fields.Item("chvFst_Name").Value)%>","<%=(rsClient.Fields.Item("intAdult_Id").Value)%>","<%=(rsClient.Fields.Item("chrSIN_no").Value)%>","<%=(rsClient.Fields.Item("chvDisability").Value)%>","<%=(rsClient.Fields.Item("dtsRefral_date").Value)%>","<%=(rsClient.Fields.Item("dtsRe_refral_date").Value)%>","","","<%=(rsClient.Fields.Item("chvStatus").Value)%>","<%=(rsClient.Fields.Item("chvCaseManager").Value)%>"<% rsClient.MoveNext(); } rsClient.Close(); %>

<%@language="JAVASCRIPT"%>
<!--#include virtual="/ASP/inc/ASPUtility.inc" -->
<!--#include virtual="/ASP/inc/ASPCheckLogin.inc" -->
<!--#include virtual="/ASP/Connections/cnnASP02.asp" -->
<!----------------------------------------------------
// Name: Daniel Chan                Date : Jun.22.2004
// Update log:
// ===========
//    - update script to suppress unique rec counter per service code
//    - retain the subtotal record count
//    - collect the unique client record count per service code     + Jun.23.2004
//---------------------------------------------------->
<%
if(String(Request.QueryString("chvFilter")) != "") { 
  rsClient__chvFilter = String(Request.QueryString("chvFilter"));
}
//----------------------------------
// list of Unique Client per service code
//----------------------------------
	var rsSummary2 = Server.CreateObject("ADODB.Recordset");
	rsSummary2.ActiveConnection = MM_cnnASP02_STRING;
	
// 	rsSummary2.Source = "{call dbo.cp_AdtClnt_SrvNote_Rpt_02A("+ Request.QueryString("inspSrtBy") + ","+ Request.QueryString("inspSrtOrd") + ",'"+ rsClient__chvFilter.replace(/'/g, "''") + "',3,0)}";
 	rsSummary2.Source = "{call dbo.cp_AdtClnt_SrvNote_Rpt_02A(11,0,'intadult_id IN (1807,1808)',3,0)}";

	rsSummary2.CursorType = 0;
	rsSummary2.CursorLocation = 2;
	rsSummary2.LockType = 3;
	rsSummary2.Open();
//----------------------------------
// Total number of Client count(duplicate) for all service group : @intAdult_Client_Rec_Cnt
// Total unique number of Client count(duplicate) for all service group : @intUnq_AdtClnt_Rec_Cnt
//----------------------------------
	var cmdSummary = Server.CreateObject("ADODB.Command");
	cmdSummary.ActiveConnection = MM_cnnASP02_STRING;
	cmdSummary.CommandText = "dbo.cp_AdtClnt_SrvNote_Rpt_03";
	cmdSummary.Parameters.Append(cmdSummary.CreateParameter("RETURN_VALUE", 3, 4));
	cmdSummary.Parameters.Append(cmdSummary.CreateParameter("@inspSrtBy", 2, 1,1,Request.QueryString("inspSrtBy")));
	cmdSummary.Parameters.Append(cmdSummary.CreateParameter("@inspSrtOrd", 2, 1,1,Request.QueryString("inspSrtOrd")));
	cmdSummary.Parameters.Append(cmdSummary.CreateParameter("@chvFilter", 200, 1,150,rsClient__chvFilter));
	cmdSummary.Parameters.Append(cmdSummary.CreateParameter("@intAdult_Client_Rec_Cnt", 3, 2));
	cmdSummary.Parameters.Append(cmdSummary.CreateParameter("@intRtnFlag", 3, 2));
	cmdSummary.Execute();
//----------------------------------
// list of service code fits the search criteria
//----------------------------------
	var rsSummary = Server.CreateObject("ADODB.Recordset");
	rsSummary.ActiveConnection = MM_cnnASP02_STRING;
	rsSummary.Source = "{call dbo.cp_adtclnt_srvnote_rpt_03("+ Request.QueryString("inspSrtBy") + ","+ Request.QueryString("inspSrtOrd") + ",'"+ rsClient__chvFilter.replace(/'/g, "''") + "',0,0)}";
	rsSummary.CursorType = 0;
	rsSummary.CursorLocation = 2;
	rsSummary.LockType = 3;
	rsSummary.Open();
//----------------------------------
// Query All Client
//----------------------------------
    var rsClient = Server.CreateObject("ADODB.Recordset");
    rsClient.ActiveConnection = MM_cnnASP02_STRING;
	rsClient.Source = "{call dbo.cp_AdtClnt_SrvNote_Rpt_02A("+ Request.QueryString("inspSrtBy") + ","+ Request.QueryString("inspSrtOrd") + ",'"+ rsClient__chvFilter.replace(/'/g, "''") + "',0,0)}";
    rsClient.CursorType = 0;
    rsClient.CursorLocation = 2;
    rsClient.LockType = 3;
    rsClient.Open();

var rsClient_numRows = 0;
var rsClient_total = 0;
while (!rsClient.EOF){
	rsClient_total++;
	rsClient.MoveNext();
}
rsClient.Requery();
%>

<html>
<head>
	<title>Client - Browse</title>
	
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>

	<Script language="Javascript">
	if (window.focus) self.focus();
	</Script>
</head>
<body>
<h3>Client - Browse</h3>
<table cellspacing="1" width="292">
  <tr> 
		<td>Displaying <b><%=rsClient_total%></b> Records</td>
    </tr>
</table>
<table class="MTable" cellpadding="2" cellspacing="1" width="960">
  <tr> 
    <th nowrap class="headrow" align="left" width="50%">Client Name</th>
    <th nowrap class="HEADROW" align="center" width="9%">EPPD Client Number</th>
    <th nowrap class="headrow" align="center" width="9%">Date of Birth</th>
    <th nowrap class="HEADROW" align="center" width="3%">SIN</th>
    <th nowrap class="headrow" align="center" width="8%">Disability 1</th>
    <th nowrap class="headrow" align="center" width="8%">Disability 2</th>
    <th nowrap class="headrow" align="center" width="8%">Institution</th>
    <th nowrap class="headrow" align="center" width="5%">Service</th>
    <th nowrap class="headrow" align="center" width="9%">Service Date</th>
    <th nowrap class="headrow" align="center" width="5%">Region</th>
    <th nowrap class="headrow" align="center" width="6%">Status</th>
    <!--
        <th nowrap class="headrow" align="left">Referral Date</th>
        <th nowrap class="headrow" align="left">Re-referral Date</th>
-->
  </tr>
  <%
// + Jun.22.2004
// Declare var.   
   var bitFlag = 0;
   var intAC_Id, chvSCDsc, intAdult_Cnt, intSrvCnt, intItemSubCnt,intTtlRCnt,intUnq_AdtCnt
   var StgTmp ; 
// init. var.
   intAdult_Cnt = 0;
   intSrvCnt    = 0; 
   intItemSubCnt= 0;
   StgTmp       = "";
   intTtlRCnt = cmdSummary.Parameters.Item("@intAdult_Client_Rec_Cnt").Value
// + Jun.23.2004
   intUnq_AdtCnt= 1;
%>
  <% 
while (!rsClient.EOF) { 
%>
  <%
    // temp buffer
    if (bitFlag == 0) {
    // very beginning
       intAC_Id = rsClient.Fields.Item("intAdult_Id").Value
       chvSCDsc = rsClient.Fields.Item("chvSrv_Code_Desc").Value
       bitFlag  = 1
       intSrvCnt++

    } else  {
       // case Service Code changes
       if (chvSCDsc != rsClient.Fields.Item("chvSrv_Code_Desc").Value ) {
           StgTmp = "<br>current Adult Count = "+intAdult_Cnt+ " and Adult ID is " +intAC_Id+" for Service Type "+chvSCDsc +"<br>";
%>
  <!--  <br>-->
  <tr> 
    <td width="50%"> 
      <!--
// + Jun.22.2004
      <p>rec Count : <%=(intAdult_Cnt)%> <br>
        ID: <%=(intAC_Id)%> <br>
        Type: <%=(chvSCDsc)%> </p>
// -->
      <br>
      <b>Service Type:</b> <%=chvSCDsc %><br>
      <b>Subtotal:</b> <%=intItemSubCnt %><br>
      <!--
//  + Jun.23.2004
//-->
      <b>Unique Client Count:</b> <%= intUnq_AdtCnt %><br>
    </td>
  </tr>
  <%
           intSrvCnt++;
           // reset var.
		   intAdult_Cnt = 0
		   intItemSubCnt= 0
           chvSCDsc = rsClient.Fields.Item("chvSrv_Code_Desc").Value
           intAC_Id = rsClient.Fields.Item("intAdult_Id").Value 
		   // + Jun.23.2004
           intUnq_AdtCnt= 1;
	   } 
       else { 
	   // same Service Code
   
       // Client Num breaks
       if (intAC_Id != rsClient.Fields.Item("intAdult_Id").Value ) {
           StgTmp = "<BR>current Adult Count = "+intAdult_Cnt+ " and Adult ID is " +intAC_Id+" for Service Type</BR>"+chvSCDsc
%>
  <!-- <br>-->
  <!--
// + Jun.22.2004
  <tr>
    <td width="14%" height="72">rec Count :<%=(intAdult_Cnt)%> <br>
      ID:<%=(intAC_Id)%> <br>
      Type:<%=(chvSCDsc)%> </td>
  </tr>
//-->
  <!--  <br>-->
  <%
           // reset var.
		   intAdult_Cnt = 0
           intAC_Id = rsClient.Fields.Item("intAdult_Id").Value
           // + Jun.23.2004
           intUnq_AdtCnt++
	   } 
       }
	}
%>
  <tr> 
    <td nowrap align="left" width="50%"><a href="javascript: openWindow('m001FS3.asp?intAdult_id=<%=(rsClient.Fields.Item("intAdult_Id").Value)%>','wQE01');"><%=(rsClient.Fields.Item("chvLst_Name").Value)%>, 
      <%=(rsClient.Fields.Item("chvFst_Name").Value)%></a></td>
    <td nowrap align="center" width="9%"><%=(rsClient.Fields.Item("chrPEN_num").Value)%>&nbsp;</td>
    <td nowrap align="center" width="9%"><%=FilterDateYearFirst(rsClient.Fields.Item("dtsBirth_date").Value)%>&nbsp;</td>
    <td nowrap align="center" width="3%"><%=FormatSIN(rsClient.Fields.Item("chrSIN_no").Value)%></td>
    <td nowrap align="center" width="8%"><%=(rsClient.Fields.Item("chvPrim_Dsbty").Value)%>&nbsp;</td>
    <td nowrap align="center" width="8%"><%=(rsClient.Fields.Item("chvSec_Dsbty").Value)%>&nbsp;</td>
    <td nowrap align="left" width="8%"><%=(rsClient.Fields.Item("chvschool").Value)%>&nbsp;</td>
    <td nowrap align="center" width="5%"><%=(rsClient.Fields.Item("chvSrv_Code_Desc").Value)%>&nbsp;</td>
    <td nowrap align="center" width="9%"><%=FilterDate(rsClient.Fields.Item("dtsService_Date").Value)%>&nbsp;</td>
    <td nowrap align="center" width="5%"><%=(rsClient.Fields.Item("chRegion_name").Value)%>&nbsp;</td>
    <td nowrap align="center" width="6%"><%=(rsClient.Fields.Item("chvCur_Status").Value)%>&nbsp;</td>
    <!--		
        <td nowrap align="center"><%=FilterDate(rsClient.Fields.Item("dtsRefral_date").Value)%>&nbsp;</td>
        <td nowrap align="center"><%=FilterDate(rsClient.Fields.Item("dtsRe_refral_date").Value)%>&nbsp;</td>
-->
  </tr>
  <%
	rsClient.MoveNext();
  // Current Adult Client counter
  intAdult_Cnt++;
  intItemSubCnt++;
}
%>
  <%
    // handle the very last entry
   if (bitFlag = 1) {
%>
  <!--       <BR><BR> -->
  <tr> 
    <td width="50%" height="77"> 
      <!--
// + Jun.22.2004
	rec Count: <%=(intAdult_Cnt)%> <BR>
       ID: <%=intAC_Id %><BR> Type: <%=chvSCDsc %>
-->
      <br>
      <b>Service Type:</b> <%=chvSCDsc %><br>
      <b>Subtotal:</b> <%=intItemSubCnt %><br>
      <!--
//  + Jun.23.2004
//-->
      <b>Unique Client Count:</b> <%= intUnq_AdtCnt %><br>
    </td>
  </tr>
  <%
   }
%>
</table>
<%
// + Jun.22.2004
//   }
%>
<BR>
<BR>
<b>Total No. of Adult record:</b> <%=intTtlRCnt %> <BR>
<b>Total Number of Service Type:</b> <%=intSrvCnt%> <BR>
<BR>    

<br>
<table cellpadding="2" cellspacing="1">
	<tr>
    <td><b>Total # of Client for all Service Code:</b></td>
		<td align="right"><%=cmdSummary.Parameters.Item("@intAdult_Client_Rec_Cnt").Value%></td>
	</tr>
<%
while (!rsSummary.EOF) {
%>
	<tr>
		<td># of <%=rsSummary.Fields.Item("chvSrv_Code_Desc").Value%></td>
		<td align="right"><%=rsSummary.Fields.Item("intRCnt").Value%></td>
	</tr>
<%
	rsSummary.MoveNext();
}
%>		
</table>

<br>
<table cellpadding="2" cellspacing="1">
	<tr>
    <td><b>Total number of unique Client for all Service Code:</b></td>
		<td align="right"><%=cmdSummary.Parameters.Item("@intUnq_AdtClnt_Rec_Cnt").Value%></td>
	</tr>
<%
while (!rsSummary2.EOF) {
%>
	<tr>
		<td># of <%=rsSummary2.Fields.Item("chvSrv_Code_Desc").Value%></td>
		<td align="right"><%=rsSummary2.Fields.Item("intUnq_AdtCnt").Value%></td>
	</tr>
<%
	rsSummary2.MoveNext();
}
%>		
</table>

</body>
</html>
<%
rsClient.Close();
%>
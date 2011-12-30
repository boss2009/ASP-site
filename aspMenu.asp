<!--------------------------------------------------------------------------
* File Name: aspMenu.asp
* Title: @WIS Master Menu
* Main SP: cp_Idv_Staff
* Description: Main menu page.  Links to all database modules and adminstration.
* Author: D. T. Chan
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="inc/ASPUtility.inc" -->
<!--#INCLUDE File="inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="Connections/cnnASP02.asp" -->
<%
var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_Idv_Staff("+ Session("insStaff_id") + ")}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>


<html>
<head>
	<title>Demo Master Menu</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	function openWindow(page){
		if (page!='nothing') win1=window.open(page, "");
		return ;
	}
	</Script>	
    <style type="text/css">
<!--
.style1 {color: #0000FF}
.style2 {
	font-family: "Times New Roman";
	font-weight: bold;
	font-size: 9pt;
}
.style3 {
	font-family: "Times New Roman", Times, serif;
	font-size: 9pt;
	font-weight: bold;
}
.style4 {
	font-family: "Times New Roman", Times, serif;
	font-size: 10pt;
	color: #000000;
}
.style5 {
	font-family: "Times New Roman", Times, serif;
	font-size: 10pt;
}
.style6 {
	font-family: "Times New Roman", Times, serif;
	font-size: 9pt;
}
.style9 {color: #006600}
.style10 {color: #006633}
.style11 {font-family: "Times New Roman", Times, serif; font-size: 9pt; color: #0000FF; }
.style12 {
	font-family: "Times New Roman";
	font-size: 9pt;
}
.style13 {color: #CC9900}
.style15 {color: #FFFFFF}
.style16 {
	font-size: 12pt;
	font-weight: bold;
}
-->
    </style>
</head>
<body onLoad="AC.focus();">
<div style="position: relative; top: 50px; centered">
<table align="center" cellspacing="2" cellpadding="5">
	<tr>
		<td colspan="3"><div align="center" class="style5 style13">
		  <div align="center">
		    <p class="style16">MASTER DEMO MENU</p>
		    <p class="style12 style15">Sirius Innovations Inc. Copyright @ 2005 - 2006 </p>
		  </div>
		</div></td>
	    <td><div align="left"><img src="i/CA.gif" alt="logo" width="68" height="50"></div></td>
	</tr>	
	<tr>
		<td width="177" align="left" bgcolor="#999999" style="font-size: 8pt; ><span class="style4">User: </span> <%=Session("insStaff_id")%> - <%=(rsStaff.Fields.Item("chvName").Value)%></td>
		<td width="181" align="left" bgcolor="#999999" style="font-size: 8pt; ><span class="style5">Permission Level: </span> <%=Session("MM_UserAuthorization")%></td>
		<td colspan="2" align="left" bgcolor="#999999" style="font-size: 8pt; ><span class="style5">Logged on:</span> <%=Session("TimeLoggedOn")%></td>
	</tr>
	<tr>
		<td colspan="4">&nbsp;</td>
	</tr>	
    <tr> 
		<td nowra><div align="center"><a href="AC/m001FS2.asp" target="Client" img="" id="AC" tabindex="1"><img src="i/tn_client_01.jpg" alt="client" width="81" height="60" align="absmiddle"></a></div></td>
		<td nowrap><div align="center"><a href="EC/m007FS2.asp" target="InventoryClass" class="style1 style3" tabindex="6"><img src="i/tn_inv_class_01.jpg" alt="Inventory Class" width="80" height="60"></a></div></td>
		<td width="174" nowrap><div align="center"><a href="LN/m008FS2.asp" target="Loan" class="style1 style2" tabindex="11"><img src="i/tn_loan_01.jpg" alt="Loan" width="80" height="60"></a></div></td>		
		<td width="117" nowrap><div align="center"><a href="FL/m032FS2.asp" target="Form" id="FL" tabindex="32"><img src="i/tn_letter.jpg" alt="Form Letters" width="80" height="60"></a></div></td>
    </tr>
    <tr>
		<td nowrap><div align="center"><a href="SH/m012FS2.asp" target="Institution" class="style1 style3" tabindex="2"><img src="i/tn_institution_01.jpg" alt="institution" width="80" height="60" border="0"></a></div></td>		
		<td nowrap><div align="center"><a href="IV/m003FS2.asp" target="Inventory" class="style1 style3" tabindex="7"><img src="i/tn_inventory_01.jpg" alt="inventory" width="80" height="60"></a></div></td>
		<td nowrap><div align="center"><a href="BO/m010FS2.asp" target="Buyout" class="style1 style3" tabindex="12"><img src="i/tn_buyout_01.jpg" alt="Buyout" width="80" height="60"></a></div></td>
		<td nowrap><div align="center"><a href="RP/m031FS2.asp" target="Report" id="RP" tabindex="31"><Img src="i/tn_report.jpg" alt="report" width="80" height="63" border="0"></a></div></td>			
    </tr>
    <tr> 
		<td nowrap><div align="center"><a href="CT/m004FS2.asp" target="Contact" class="style1 style3" tabindex="3"><img src="i/tn_CONTACT_02.jpg" alt="contact" width="80" height="60"></a></div></td>		
		<td nowrap><div align="center"><a href="PR/m014FS2.asp" target="Requisition" class="style1 style3" tabindex="8"><img src="i/tn_pur_req_02.jpg" alt="Purchase Requisition" width="80" height="60"></a></div></td>
		<td nowrap><div align="center"><a href="ES/m009FS2.asp" target="EquipmentService" class="style1 style3" tabindex="13"><img src="i/tn_service_01.jpg" alt="Equipment Service" width="80" height="60"></a></div></td>
		<td nowrap><div align="center"><a href="IC/m033FS2.asp"><img src="i/tn_invoice.jpg" alt="Invoices" width="80" height="66" border="0"></a></div></td>		
    </tr>
    <tr> 		
		<td nowrap><div align="center"><a href="CP/m006FS2.asp" target="Organization" class="style1 style3" tabindex="4"><img src="i/tn_organization_02.jpg" alt="Organization" width="81" height="60" ></a></div></td>				
		<td nowrap><div align="center"><a href="BD/m005FS2.asp" target="Bundle" class="style1 style3" tabindex="9"><img src="i/tn_equip_bundle_01.jpg" alt="Equipment Bundle" width="80" height="60"></a></div></td>		
		<td></td>
		<td><div align="center"><a href="OL/m034FS2.asp"><Img src="i/tn_data_warehse.jpg" alt="OLAP data analysis" width="81" height="69" border="0"></a></div></td>
    </tr>
	<tr>
		<td nowrap><div align="center"><a href="ST/m002FS2.asp" target="Staff" class="style1 style3" tabindex="5"><img src="i/tn_staff_01.jpg" alt="staff" width="81" height="60"></a></div></td>				
		<td nowrap><div align="center"><a href="PS/m022FS2.asp" target="PILATStudent" class="style1 style3" tabindex="10"><img src="i/tn_student_01.jpg" alt="Student" width="80" height="60"></a></div></td>
		<td></td>
		<td></td>
	</tr>
	<tr>
		<td colspan="4">&nbsp;</td>
	</tr>
	<tr bgcolor="#999999">
		<td align="left" style="font-size: 7pt; font-weight: bold"><div align="center">
	      <%if (Session("MM_UserAuthorization") >= 5 ){%>
		      <a href="Admin/m018Menu.asp" target="Admin" class="style6 style9"><img src="i/tn_admin_01.jpg" alt="Adminsitration" width="80" height="60"></a>
          <%}%>
		  </div></td>
		<td align="left" style="font-size: 7pt; font-weight: bold"><div align="center">
	      <%if (Session("MM_UserAuthorization") >= 5 ){%>
		      <a href="PL/m020FS2.asp" target="Maintenance" class="style6 style10"><img src="i/tn_main_tool_01.jpg" alt="Maintenance Tools" width="80" height="60"></a>
          <%}%>
		  </div></td>
		<td></td>
		<td align="left" style="font-size: 7pt; font-weight: bold"><a href="asplogout.asp" class="style11">Log Out</a></td>
	</tr>
</table>
</div>
</body>
</html>
<%
rsStaff.Close();
%>
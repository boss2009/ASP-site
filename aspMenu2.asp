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
rsStaff.Source = "{call dbo.cp_Idv_Staff("+ Session("insStaff_id") + ",0)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title> Master Menu</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	function openWindow(page){
		if (page!='nothing') win1=window.open(page, "");
		return ;
	}
	</Script>	
</head>
<body onLoad="AC.focus();">
<div style="position: relative; top: 50px; centered">
<table align="center" cellspacing="2" cellpadding="5">
	<tr>
		<td align="left" style="font-size: 7pt; font-weight: bold">User: <%=(rsStaff.Fields.Item("chvName").Value)%></td>
		<td align="left" style="font-size: 7pt; font-weight: bold">Permission Level: <%=Session("MM_UserAuthorization")%></td>
		<td align="left" style="font-size: 7pt; font-weight: bold" colspan="2">Logged on: <%=Session("TimeLoggedOn")%></td>
	</tr>
	<tr>
		<td colspan="4"><hr></td>
	</tr>	
    <tr> 
		<td nowra><a href="AC/m001FS2.asp" tabindex="1" id="AC" target="Client">Client</a></td>
		<td nowrap><a href="EC/m007FS2.asp" tabindex="6" target="InventoryClass">Inventory Class</a></td>
		<td nowrap><a href="LN/m008FS2.asp" tabindex="11" target="Loan">Loan</a></td>		
		<td nowrap>Form Letters</td>				
    </tr>
    <tr>
		<td nowrap><a href="SH/m012FS2.asp" tabindex="2" target="Institution">Institution</a></td>		
		<td nowrap><a href="IV/m003FS2.asp" tabindex="7" target="Inventory">Inventory</a></td>
		<td nowrap><a href="BO/m010FS2.asp" tabindex="12" target="Buyout">Buyout</a></td>
		<td nowrap>Forms &amp; Reports</td>			
    </tr>
    <tr> 
		<td nowrap><a href="CT/m004FS2.asp" tabindex="3" target="Contact">Contact</a></td>		
		<td nowrap><a href="PR/m014FS2.asp" tabindex="8" target="Requisition">Purchase Requisition</a></td>
		<td nowrap><a href="ES/m009FS2.asp" tabindex="13" target="EquipmentService">Equipment Service</a></td>
		<td nowrap>Invoice</td>		
    </tr>
    <tr> 		
		<td nowrap><a href="CP/m006FS2.asp" tabindex="4" target="Organization">Organization</a></td>				
		<td nowrap><a href="BD/m005FS2.asp" tabindex="9" target="Bundle">Equipment Bundle</a></td>		
		<td></td>
		<td></td>
    </tr>
	<tr>
		<td nowrap><a href="ST/m002FS2.asp" tabindex="5" target="Staff">Staff</a></td>				
		<td nowrap><a href="PS/m022FS2.asp" tabindex="10" target="PILATStudent">PILAT Student</a></td>
		<td></td>
		<td></td>
	</tr>
	<tr>
		<td colspan="4"><hr></td>
	</tr>
	<tr>
		<td align="left" style="font-size: 7pt; font-weight: bold"><%if (Session("MM_UserAuthorization") >= 5 ){%><a href="Admin/m018Menu.asp" target="Admin">Administration</a><%}%></td>
		<td align="left" style="font-size: 7pt; font-weight: bold"><%if (Session("MM_UserAuthorization") >= 5 ){%><a href="PL/m020FS2.asp" target="Maintenance">Maintenance Tools</a><%}%></td>
		<td></td>
		<td align="left" style="font-size: 7pt; font-weight: bold"><a href="asplogout.asp">Log Out</a></td>
	</tr>
</table>
</div>
</body>
</html>
<%
rsStaff.Close();
%>
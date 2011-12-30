<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title>Select Staff</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/MyFunctions.js"></script>	
	<script language="JavaScript">
	function Init() {
		document.frm08p02.StaffName.focus();
	}

	function SelectStaff(){
		opener.document.frm0101.UserType.value=1;
		opener.document.frm0101.IndividualUserID.value=document.frm08p02.StaffName[document.frm08p02.StaffName.selectedIndex].value;
		opener.document.frm0101.IndividualUserName.value=document.frm08p02.StaffName.options[document.frm08p02.StaffName.selectedIndex].text;
		self.close();
	}
	</script>
</head>
<body onload="Init();">
<form name="frm08p02" method="post" action="">
<h5>Select Staff</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><select name="StaffName" tabindex="1" accesskey="F">
			<%
			while (!rsStaff.EOF) {
			%>
				<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>"><%=rsStaff.Fields.Item("chvName").Value%> 
			<%
				rsStaff.MoveNext();
			}
			rsStaff.MoveFirst();
			%>		
		</select></td>		
		<td nowrap>
			<input type="button" value="Select Staff" tabindex="2" onClick="SelectStaff();" class="btnstyle">
			<input type="button" value="Close" tabindex="3" onClick="top.window.close();" class="btnstyle">
		</td>
	</tr>
</table>
<input type="hidden" name="MM_flag" value="false">
</form>
</body>
</html>
<%
rsStaff.Close();
%>
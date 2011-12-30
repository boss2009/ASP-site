<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsVendor = Server.CreateObject("ADODB.Recordset");
rsVendor.ActiveConnection = MM_cnnASP02_STRING;
rsVendor.Source = "{call dbo.cp_Get_EqCls_Vendor("+Request.QueryString("ClassID")+",0,0)}";
rsVendor.CursorType = 0;
rsVendor.CursorLocation = 2;
rsVendor.LockType = 3;
rsVendor.Open();
%>
<html>
<head>
	<title>Vendors</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				document.frm0103c.submit();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=700,height=400,scrollbars=1,left=0,top=0,status=1");
		return ;
	}			
	</script>
</head>
<body>
<h5>Vendors</h5>
<hr>
<table cellspacing="1" cellpadding="1">
	<tr> 
		<th class="headrow" align="left" nowrap>&nbsp;</th>
		<th class="headrow" align="left" nowrap>Vendor</th>
		<th class="headrow" align="left" nowrap>List Unit Cost</th>
		<th class="headrow" align="left" nowrap>Contract PO</th>		
		<th class="headrow" align="left" nowrap>Entry Date</th>
		<th class="headrow" align="left" nowrap>Address</th>
		<th class="headrow" align="left" nowrap>Province/State</th>
		<th class="headrow" align="left" nowrap>Phone Number</th>
	</tr>
<% 
while (!rsVendor.EOF) { 
%>
	<tr> 
<%	
	if (rsVendor.Fields.Item("bitIsCurrent").Value == "1") { 	
%>
		<td nowrap>Default</td>	
<%
	} else {
%>		
		<td nowrap>&nbsp;</td>		
<%		
	}		
%>		
		<td nowrap width="200"><a href="m007e0103d.asp?ClassID=<%=Request.QueryString("ClassID")%>&intEqCls_Dtl_id=<%=(rsVendor.Fields.Item("intEqCls_Dtl_id").Value)%>"><%=(rsVendor.Fields.Item("chvCompany_Name").Value)%></a></td>
		<td nowrap align="right"><%=FormatCurrency(rsVendor.Fields.Item("fltList_Unit_Cost").Value)%></td>
		<td nowrap align="left"><%=(rsVendor.Fields.Item("chvContract_PO").Value)%></td>		
		<td nowrap align="center"><%=FilterDate(rsVendor.Fields.Item("dtsEntry_Date").Value)%></td>
		<td nowrap align="left"><%=(rsVendor.Fields.Item("chvAddress").Value)%></td>
		<td nowrap align="center"><%=(rsVendor.Fields.Item("chrprvst_abbv").Value)%></td>
		<td nowrap align="left"><%=FormatPhoneNumber(rsVendor.Fields.Item("chvPhone_Type").Value,rsVendor.Fields.Item("chvPhone1_Arcd").Value,rsVendor.Fields.Item("chvPhone1_Num").Value,rsVendor.Fields.Item("chvPhone1_Ext").Value,"","","","","","","","","")%>&nbsp;</td>
    </tr>		
<% 
	rsVendor.MoveNext();
} 
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><a href="javascript: openWindow('m007a0103d.asp?ClassID=<%=Request.QueryString("classid")%>','wA0103c');">Add Vendor</a></td>
	</tr>
</table>
<input type="hidden" name="ClassID" value="<%=Request.QueryString("classid")%>">
</body>
</html>
<%
rsVendor.Close();
%>
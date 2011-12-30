<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_actionAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_actionAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_action"))=="delete"){
	var rsDeleteClass = Server.CreateObject("ADODB.Recordset");
	rsDeleteClass.ActiveConnection = MM_cnnASP02_STRING;
	rsDeleteClass.Source = "{call dbo.cp_Delete_Eqp_Class2(" + Request.QueryString("ClassID") + ","+Request.Form("ClassDetailID")+",0)}";	
	rsDeleteClass.CursorType = 0;
	rsDeleteClass.CursorLocation = 2;
	rsDeleteClass.LockType = 3;
	rsDeleteClass.Open();
	Response.Redirect("AddDeleteSuccessful3.asp?action=Deleted");	
}

if (String(Request.Form("MM_action"))=="update"){
	var ConcreteClassName = String(Request.Form("ConcreteClassName")).replace(/'/g, "''");		
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");		
	var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
	rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
	rsSubAbstractClass.Source = "{call dbo.cp_Update_Eqp_Class(" + Request.Form("ClassID") + ",'" + ConcreteClassName + "'," + Request.Form("SubAbstractClassID") + "," + Request.Form("ClassStatus") +"," + Session("insStaff_id") + ",'" + Request.Form("SubjectTo") +"',0,0,0,0,'" + Request.Form("ModelNumber") + "','" + Notes + "','C',0)}";
	rsSubAbstractClass.CursorType = 0;
	rsSubAbstractClass.CursorLocation = 2;
	rsSubAbstractClass.LockType = 3;
	rsSubAbstractClass.Open();
	Response.Redirect("UpdateSuccessful2.asp?page=m007e0103.asp&ClassID="+Request.Form("ClassID"));
}

var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW(" + Request.QueryString("ClassID") + ",'C',1)}";	
rsConcreteClass.CursorType = 0;
rsConcreteClass.CursorLocation = 2;
rsConcreteClass.LockType = 3;
rsConcreteClass.Open();	

var rsWarrantyLength = Server.CreateObject("ADODB.Recordset");
rsWarrantyLength.ActiveConnection = MM_cnnASP02_STRING;
rsWarrantyLength.Source = "{call dbo.cp_ASP_lkup(62)}";
rsWarrantyLength.CursorType = 0;
rsWarrantyLength.CursorLocation = 2;
rsWarrantyLength.LockType = 3;
rsWarrantyLength.Open();
%>
<html>
<head>
	<title>Concrete Class</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
	   		case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		if (!CheckTextArea(document.frm0103.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		if (Trim(document.frm0103.ConcreteClassName.value)==""){
			alert("Enter Concrete Class Name.");
			document.frm0103.ConcreteClassName.focus();
			return ;
		}
		document.frm0103.submit();
	}
	
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=500,height=400,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	
	function popVendorDeal(){
		openWindow("m007pop.asp?ClassID=<%=rsConcreteClass.Fields.Item("insCnctCls_id").Value%>&intEqCls_dtl_id=<%=rsConcreteClass.Fields.Item("intEqCls_Dtl_id").Value%>","");
	}
	
	function DeleteClass(){
		if (confirm("Delete This Class?")) {
			document.frm0103.MM_action.value="delete";
			document.frm0103.submit();
		} 		
	}
	
	function Init(){
	<%
	if (!((rsConcreteClass.Fields.item("insVendor_id").Value > 0) && (rsConcreteClass.Fields.item("bitIsCurrent").Value == 1))){	
	%>
		document.frm0103.btnVendorDeal.disabled = true;
	<%
	}
	%>
		document.frm0103.AbstractClassName.focus();
	}		
	</script>
</head>
<body onLoad="document.frm0103.AbstractClassName.focus();"> 
<form action="<%=MM_actionAction%>" method="POST" name="frm0103">
<h5>Concrete Class</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Abstract Class Name:</td>
		<td nowrap><input type="text" name="AbstractClassName" maxlength="50" value="<%=rsConcreteClass.Fields.Item("chvAbsClsName").Value%>" size="50" readonly tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td nowrap>Sub Abstract Class Name:</td>
		<td nowrap><input type="text" name="SubAbstractClassName" maxlength="50" value="<%=rsConcreteClass.Fields.Item("chvSubAbsClsName").Value%>" size="50" readonly tabindex="2"></td>
	<tr>
		<td nowrap>Concrete Class Name:</td>
		<td nowrap><input type="text" name="ConcreteClassName" maxlength="50" value="<%=rsConcreteClass.Fields.Item("chvName").Value%>" size="50" tabindex="3"></td>
	</tr>
	<tr>
		<td nowrap>Model Number:</td>
		<td nowrap><input type="text" name="ModelNumber" value="<%=(rsConcreteClass.Fields.Item("chvModel_Number").Value)%>" maxlength="50" size="15" tabindex="4"></td>		
	</tr>
	<tr>
		<td nowrap>List Unit Cost:</td>
		<td nowrap>
			<input type="text" name="ListUnitCost" value="<%=FormatCurrency(rsConcreteClass.Fields.Item("fltList_Unit_Cost").Value)%>" size="15" tabindex="5" readonly>
			<input type="button" name="btnVendorDeal" value="Change" onClick="popVendorDeal();" class="btnstyle">
		</td>
	</tr>	
	<tr>
		<td nowrap>Subject To:</td>
		<td nowrap><select name=SubjectTo tabindex="6">
			<option value="0" <%=((rsConcreteClass.Fields.Item("chvSbjTotax").Value == "0")?"SELECTED":"")%>>No Tax
			<option value="1" <%=((rsConcreteClass.Fields.Item("chvSbjTotax").Value == "1")?"SELECTED":"")%>>PST
			<option value="2" <%=((rsConcreteClass.Fields.Item("chvSbjTotax").Value == "2")?"SELECTED":"")%>>GST
			<option value="3" <%=((rsConcreteClass.Fields.Item("chvSbjTotax").Value == "3")?"SELECTED":"")%>>PST/GST
		</select></td>
	</tr>	
	<tr> 
		<td nowrap>Parts Warranty Length:</td>
		<td nowrap><select name="PartsWarrantyLength" tabindex="7" disabled>
		<% 
		while (!rsWarrantyLength.EOF) { 
		%>
			<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" <%=((rsConcreteClass.Fields.Item("insPartsWLen").Value==rsWarrantyLength.Fields.Item("insWarrenty_id").Value)?"SELECTED":"")%>><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%>
		<% 
		rsWarrantyLength.MoveNext();			
		} 
		%>
		</select></td>
	</tr>	
	<tr> 
		<td nowrap>Labour Warranty Length:</td>
		<td nowrap><select name="LabourWarrantyLength" tabindex="8" disabled>	  
		<% 
		rsWarrantyLength.MoveFirst();			
		while (!rsWarrantyLength.EOF) { 			
		%>
			<option value="<%=(rsWarrantyLength.Fields.Item("insWarrenty_id").Value)%>" <%=((rsConcreteClass.Fields.Item("insLaborWLen").Value==rsWarrantyLength.Fields.Item("insWarrenty_id").Value)?"SELECTED":"")%>><%=(rsWarrantyLength.Fields.Item("chvWarrenty_Dsc").Value)%>
		<% 
		rsWarrantyLength.MoveNext();
		} 
		%>
		</select></td>
	</tr>	
	<tr>
		<td nowrap>Class Status:</td>
		<td nowrap><select name=ClassStatus tabindex="9">
			<option value="1" <%=((rsConcreteClass.Fields.Item("bitIs_Class_Active").Value == "1")?"SELECTED":"")%>>Active
			<option value="0" <%=((rsConcreteClass.Fields.Item("bitIs_Class_Active").Value == "0")?"SELECTED":"")%>>Inactive
		</select></td>
	</tr>	
	<tr>
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" tabindex="10" rows="5" cols="59" accesskey="L"><%=rsConcreteClass.Fields.Item("chvComments_specs").Value%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="11" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="12" class="btnstyle"></td>
<% 
	if (Session("MM_UserAuthorization") >= 5) { 
%>
		<td><input type="button" value="Delete" onClick="DeleteClass();" tabindex="13" class="btnstyle"></td>
<% 
	} 
%>		
	</tr>
</table>
<input type="hidden" name="SubAbstractClassID" value="<%=rsConcreteClass.Fields.Item("insSubAbsCls_id").Value%>">
<input type="hidden" name="ClassID" value="<%=Request.QueryString("ClassID")%>">
<input type="hidden" name="ClassDetailID" value="<%=rsConcreteClass.Fields.Item("intEqCls_Dtl_id").Value%>">
<input type="hidden" name="MM_action" value="update">
</form>
</body>
</html>
<%
rsConcreteClass.Close();
rsWarrantyLength.Close();
%>
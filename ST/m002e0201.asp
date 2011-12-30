<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_action")) == "update") {
	var rsStaffAddress = Server.CreateObject("ADODB.Recordset");
	rsStaffAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsStaffAddress.Source = "{call dbo.cp_staff_address(0,"+Request.Form("MM_recordId")+",'"+String(Request.Form("StreetAddress")).replace(/'/g, "''")+"','"+String(Request.Form("City")).replace(/'/g, "''")+"',"+Request.Form("ProvinceState")+",'"+Trim(Request.Form("PostalCode"))+"',"+Request.Form("PrimaryPhoneType")+",'"+Trim(Request.Form("PrimaryPhoneAreaCode"))+"','"+Trim(Request.Form("PrimaryPhoneNumber"))+"','"+Trim(Request.Form("PrimaryPhoneExtension"))+"',"+Request.Form("SecondaryPhoneType")+",'"+Request.Form("SecondaryPhoneAreaCode")+"','"+Trim(Request.Form("SecondaryPhoneNumber"))+"','"+Trim(Request.Form("SecondaryPhoneExtension"))+"',0,'','','','"+Request.Form("Email")+"','',9,0,'E',0)}";
	rsStaffAddress.CursorType = 0;
	rsStaffAddress.CursorLocation = 2;
	rsStaffAddress.LockType = 3;
	rsStaffAddress.Open();
	Response.Redirect("m002e0201.asp?insStaff_id="+Request.QueryString("insStaff_id"));
}

if (String(Request("MM_action")) == "insert") {
	var rsStaffAddress = Server.CreateObject("ADODB.Recordset");
	rsStaffAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsStaffAddress.Source = "{call dbo.cp_staff_address("+Request.QueryString("insStaff_id")+",0,'"+String(Request.Form("StreetAddress")).replace(/'/g, "''")+"','"+String(Request.Form("City")).replace(/'/g, "''")+"',"+Request.Form("ProvinceState")+",'"+Request.Form("PostalCode")+"',"+Request.Form("PrimaryPhoneType")+",'"+Request.Form("PrimaryPhoneAreaCode")+"','"+Request.Form("PrimaryPhoneNumber")+"','"+Request.Form("PrimaryPhoneExtension")+"',"+Request.Form("SecondaryPhoneType")+",'"+Request.Form("SecondaryPhoneAreaCode")+"','"+Request.Form("SecondaryPhoneNumber")+"','"+Request.Form("SecondaryPhoneExtension")+"',0,'','','','"+Request.Form("Email")+"','',9,0,'A',0)}";
	rsStaffAddress.CursorType = 0;
	rsStaffAddress.CursorLocation = 2;
	rsStaffAddress.LockType = 3;
	rsStaffAddress.Open();
	Response.Redirect("m002e0201.asp?insStaff_id="+Request.QueryString("insStaff_id"));
}

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_staff2("+Request.QueryString("insStaff_id")+",0,'','',0,'','',0,0,0,0,0,0,0,0,0,1,0,'',1,'Q',0)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var intAddress_id = 0;
var IsNew = false;
if (rsStaff.EOF) {
  IsNew = true;
} else {
	if ((rsStaff.Fields.Item("intAddress_id").Value != null) && (rsStaff.Fields.Item("intAddress_id").Value != 0)) {
		intAddress_id = rsStaff.Fields.Item("intAddress_id").Value;
	} else {
		IsNew = true;
	}
}

var rsStaffAddress = Server.CreateObject("ADODB.Recordset");
rsStaffAddress.ActiveConnection = MM_cnnASP02_STRING;
rsStaffAddress.Source = "{call dbo.cp_staff_address(0,"+intAddress_id+",'','',0,'',0,'','','',0,'','','',0,'','','','','',0,1,'Q',0)}";
rsStaffAddress.CursorType = 0;
rsStaffAddress.CursorLocation = 2;
rsStaffAddress.LockType = 3;
rsStaffAddress.Open();

var rsProvince = Server.CreateObject("ADODB.Recordset");
rsProvince.ActiveConnection = MM_cnnASP02_STRING;
rsProvince.Source = "{call dbo.cp_Prov_State}";
rsProvince.CursorType = 0;
rsProvince.CursorLocation = 2;
rsProvince.LockType = 3;
rsProvince.Open();

var rsPhoneType = Server.CreateObject("ADODB.Recordset");
rsPhoneType.ActiveConnection = MM_cnnASP02_STRING;
rsPhoneType.Source = "{call dbo.cp_Phone_Type}";
rsPhoneType.CursorType = 0;
rsPhoneType.CursorLocation = 2;
rsPhoneType.LockType = 3;
rsPhoneType.Open();

var rsAreaCode = Server.CreateObject("ADODB.Recordset");
rsAreaCode.ActiveConnection = MM_cnnASP02_STRING;
rsAreaCode.Source = "{call dbo.cp_area_code(0,'',0,2,'Q',0)}";
rsAreaCode.CursorType = 0;
rsAreaCode.CursorLocation = 2;
rsAreaCode.LockType = 3;
rsAreaCode.Open();
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>Update Address</title>
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
			case 85:
				//alert("U");
				document.frm0201.reset();
			break;
		   	case 76 :
				//alert("L");
				window.location.href='m002e0101.asp?insStaff_id=<%=Request.QueryString("insStaff_id")%>';
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init(){
		document.frm0201.StreetAddress.focus();
	}
	
	function Save(){
		if (!CheckPostalCode(document.frm0201.PostalCode.value)){
			alert("Invalid Postal Code.");
			document.frm0201.PostalCode.focus();
			return ;
		}
		if (!CheckEmail(document.frm0201.EMail.value)){
			alert("Invalid Email.");
			document.frm0201.EMail.focus();
			return ;
		}
		var tempPC = document.frm0201.PostalCode.value;
		tempPC = tempPC.toUpperCase();
		document.frm0201.PostalCode.value = tempPC;		
				
		document.frm0201.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0201">
<h5>Address</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap valign="top">Street Address:</td>
		<td nowrap valign="top"><textarea name="StreetAddress" cols="30" rows="3" tabindex="1" accesskey="F"><%=((!IsNew)?rsStaffAddress.Fields.Item("chvAddress").Value:"")%></textarea></td>
	</tr>
	<tr> 
		<td nowrap>City:</td>
		<td nowrap><input type="text" name="City" value="<%=((!IsNew)?rsStaffAddress.Fields.Item("chvCity").Value:"")%>" maxlength="50" tabindex="2" ></td>
	</tr>
	<tr> 
		<td nowrap>Province/State:</td>
		<td nowrap><select name="ProvinceState" tabindex="3">
		<% 
		while (!rsProvince.EOF) {
		%>
			<option value="<%=(rsProvince.Fields.Item("intprvst_id").Value)%>" <%if (!IsNew) Response.Write(((rsProvince.Fields.Item("intprvst_id").Value==rsStaffAddress.Fields.Item("intprvst_id").Value)?"SELECTED":""))%>><%=(rsProvince.Fields.Item("chrprvst_abbv").Value)%></option>
		<%
			rsProvince.MoveNext();
		}
		%>
        </select></td>
    </tr>	
	<tr>
		<td nowrap>Country:</td>
		<td nowrap><input type="text" name="Country" value="<%=((!IsNew)?rsStaffAddress.Fields.Item("chvcntry_name").Value:"")%>" tabindex="4" readonly></td>
	</tr>
    <tr> 
		<td nowrap>Postal Code:</td>
		<td nowrap><input type="text" name="PostalCode" value="<%=((!IsNew)?FormatPostalCode(rsStaffAddress.Fields.Item("chvPostal_zip").Value):"")%>" tabindex="5" size="10" maxlength="7" onChange="FormatPostalCode(this);"></td>
    </tr>
    <tr> 
		<td nowrap>Primary Phone:</td>
		<td nowrap> 
			<select name="PrimaryPhoneType" tabindex="6">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%if (!IsNew) Response.Write(((rsPhoneType.Fields.Item("intPhone_type_id").Value==rsStaffAddress.Fields.Item("intPhone_Type_1").Value)?"SELECTED":""))%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			rsPhoneType.MoveFirst();
			%>
			</select>
			<select name="PrimaryPhoneAreaCode" tabindex="7">
				<option value="" <%if (!IsNew) Response.Write(((rsStaffAddress.Fields.Item("chvPhone1_Arcd").Value=="")?"SELECTED":""))%>>			
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%if (!IsNew) Response.Write(((rsAreaCode.Fields.Item("chvAC_num").Value==rsStaffAddress.Fields.Item("chvPhone1_Arcd").Value)?"SELECTED":""))%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			rsAreaCode.MoveFirst();
			%>			
			</select>
			<input type="text" name="PrimaryPhoneNumber" value="<%=((!IsNew)?FormatPhoneNumberOnly(rsStaffAddress.Fields.Item("chvPhone1_Num").Value):"")%>" size="9" tabindex="8" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">Ext
			<input type="text" name="PrimaryPhoneExtension" value="<%=((!IsNew)?rsStaffAddress.Fields.Item("chvPhone1_Ext").Value:"")%>" size="4" tabindex="9" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>Secondary Phone:</td>
		<td nowrap>
			<select name="SecondaryPhoneType" tabindex="10">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%if (!IsNew) Response.Write(((rsPhoneType.Fields.Item("intPhone_type_id").Value==rsStaffAddress.Fields.Item("intPhone_Type_2").Value)?"SELECTED":""))%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			%>
			</select>
			<select name="SecondaryPhoneAreaCode" tabindex="11">
				<option value="" <%if (!IsNew) Response.Write(((rsStaffAddress.Fields.Item("chvPhone2_Arcd").Value=="")?"SELECTED":""))%>>
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%if (!IsNew) Response.Write(((rsAreaCode.Fields.Item("chvAC_num").Value==rsStaffAddress.Fields.Item("chvPhone2_Arcd").Value)?"SELECTED":""))%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>			
			</select>
			<input type="text" name="SecondaryPhoneNumber" value="<%=((!IsNew)?FormatPhoneNumberOnly(rsStaffAddress.Fields.Item("chvPhone2_Num").Value):"")%>" size="9" tabindex="12" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">Ext 
			<input type="text" name="SecondaryPhoneExtension" value="<%=((!IsNew)?rsStaffAddress.Fields.Item("chvPhone2_Ext").Value:"")%>" size="4" tabindex="13" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>E-Mail:</td>
		<td nowrap><input type="text" name="EMail" value="<%=((!IsNew)?Trim(rsStaffAddress.Fields.Item("chvEmail").Value):"")%>" tabindex="14" accesskey="L"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="16" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="17" onClick="window.location.href='m002e0101.asp?insStaff_id=<%=Request.QueryString("insStaff_id")%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="<%=((IsNew)?"insert":"update")%>">
<input type="hidden" name="MM_recordId" value="<%=intAddress_id%>">
</form>
</body>
</html>
<%
rsStaffAddress.Close();
rsProvince.Close();
rsPhoneType.Close();
rsAreaCode.Close();
%>
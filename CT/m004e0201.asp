<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update")) == "true"){
	var rsAddress = Server.CreateObject("ADODB.Recordset");
	rsAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsAddress.Source="{call dbo.cp_Contact_Address(0,"+Request.Form("MM_recordId")+",'','"+String(Request.Form("StreetAddress")).replace(/'/g, "''")+"','"+String(Request.Form("City")).replace(/'/g, "''") + "',"+ Request.Form("Province") + ",'"+ Request.Form("PostalCode") + "',"+ Request.Form("PrimaryPhoneType") + ",'"+ Request.Form("PrimaryPhoneAreaCode") + "','"+ Request.Form("PrimaryPhoneNumber") + "','"+ Request.Form("PrimaryPhoneExtension") + "',"+ Request.Form("SecondaryPhoneType") + ",'"+ Request.Form("SecondaryPhoneAreaCode")+ "','"+Request.Form("SecondaryPhoneNumber")+"','"+Request.Form("SecondaryPhoneExtension") + "',0,'','','','"+ Request.Form("EMail") + "','"+ String(Request.Form("Notes")).replace(/'/g, "''") + "',0,0,'E',0)}";
	rsAddress.CursorType = 0;
	rsAddress.CursorLocation = 2;
	rsAddress.LockType = 3;
	rsAddress.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m004q0201.asp&intContact_id="+Request.QueryString("intContact_id"));
}

var rsAddress = Server.CreateObject("ADODB.Recordset");
rsAddress.ActiveConnection = MM_cnnASP02_STRING;
rsAddress.Source = "{call dbo.cp_Contact_Address(0,"+ Request.QueryString("intaddr_id") +",'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,1,'Q',0)}";
rsAddress.CursorType = 0;
rsAddress.CursorLocation = 2;
rsAddress.LockType = 3;
rsAddress.Open();

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
				history.back();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (!CheckTextArea(document.frm0201.Notes, 50)){
			alert("Text area cannot exceed 50 characters.");
			return ;
		}
	
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
<body onLoad="javascript:document.frm0201.StreetAddress.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0201">
<h5>Address</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap valign="top">Street Address:</td>
		<td nowrap valign="top"><textarea name="StreetAddress" cols="30" rows="3" tabindex="1" accesskey="F"><%=Trim(rsAddress.Fields.Item("chvAddress").Value)%></textarea></td>
	</tr>
	<tr> 
		<td nowrap>City:</td>
		<td nowrap><input type="text" name="City" value="<%=Trim(rsAddress.Fields.Item("chvCity").Value)%>" maxlength="50" tabindex="2"></td>
	</tr>
	<tr> 
		<td nowrap>Province:</td>
		<td nowrap><select name="Province" tabindex="3">
			<% 
			while (!rsProvince.EOF) {
			%>
				<option value="<%=(rsProvince.Fields.Item("intprvst_id").Value)%>" <%=((rsProvince.Fields.Item("intprvst_id").Value == rsAddress.Fields.Item("intprvst_id").Value)?"SELECTED":"")%>><%=(rsProvince.Fields.Item("chrprvst_abbv").Value)%></option>
			<%
				rsProvince.MoveNext();
			}
			%>
        </select></td>
    </tr>	
    <tr> 
		<td nowrap>Postal Code:</td>
		<td nowrap><input type="text" name="PostalCode" value="<%=FormatPostalCode(rsAddress.Fields.Item("chvPostal_zip").Value)%>" tabindex="4" size="10" maxlength="7" onChange="FormatPostalCode(this);"></td>
    </tr>
    <tr> 
		<td nowrap>Primary Phone:</td>
		<td nowrap> 
			<select name="PrimaryPhoneType" tabindex="5">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%=((rsPhoneType.Fields.Item("intPhone_type_id").Value == rsAddress.Fields.Item("intPhone_Type_1").Value)?"SELECTED":"")%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			rsPhoneType.MoveFirst();
			%>
			</select>
			<select name="PrimaryPhoneAreaCode" tabindex="6">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%=((rsAddress.Fields.Item("chvPhone1_Arcd").Value == rsAreaCode.Fields.Item("chvAC_num").Value)?"SELECTED":"")%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			rsAreaCode.MoveFirst();
			%>			
			</select>
			<input type="text" name="PrimaryPhoneNumber" value="<%=FormatPhoneNumberOnly(rsAddress.Fields.Item("chvPhone1_Num").Value)%>" size="9" tabindex="7" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">
			Ext <input type="text" name="PrimaryPhoneExtension" value="<%=Trim(rsAddress.Fields.Item("chvPhone1_Ext").Value)%>" size="4" tabindex="8" onKeypress="AllowNumericOnly();">
		</td>
    </tr>
    <tr> 
		<td nowrap>Secondary Phone:</td>
		<td nowrap>
			<select name="SecondaryPhoneType" tabindex="9">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%=((rsPhoneType.Fields.Item("intPhone_type_id").Value == rsAddress.Fields.Item("intPhone_Type_2").Value)?"SELECTED":"")%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			%>
			</select>
			<select name="SecondaryPhoneAreaCode" tabindex="10">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%=((rsAddress.Fields.Item("chvPhone2_Arcd").Value == rsAreaCode.Fields.Item("chvAC_num").Value)?"SELECTED":"")%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>			
			</select>
			<input type="text" name="SecondaryPhoneNumber" value="<%=FormatPhoneNumberOnly(rsAddress.Fields.Item("chvPhone2_Num").Value)%>" size="9" tabindex="11" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">
			Ext <input type="text" name="SecondaryPhoneExtension" value="<%=Trim(rsAddress.Fields.Item("chvPhone2_Ext").Value)%>" size="4" tabindex="12" onKeypress="AllowNumericOnly();">
		</td>
    </tr>
    <tr> 
		<td nowrap>E-Mail:</td>
		<td nowrap><input type="text" name="EMail" value="<%=Trim(rsAddress.Fields.Item("chvEmail").Value)%>" tabindex="13"></td>
	</tr>  
	<tr>
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="14" accesskey="L"><%=Trim(rsAddress.Fields.Item("chvNote").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="16" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="17" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsAddress.Fields.Item("intAddress_id").Value %>">
</form>
</body>
</html>
<%
rsAddress.Close();
rsProvince.Close();
rsPhoneType.Close();
rsAreaCode.Close();
%>
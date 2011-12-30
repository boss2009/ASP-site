<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_action"))=="update") {
	var UserName = String(Request.Form("UserName")).replace(/'/g, "''");
	var ContactFirstName = String(Request.Form("ContactFirstName")).replace(/'/g, "''");
	var ContactLastName = String(Request.Form("ContactLastName")).replace(/'/g, "''");
	var rsShippingAddress = Server.CreateObject("ADODB.Recordset");
	rsShippingAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsShippingAddress.Source = "{call dbo.cp_eqpsrv_ship_addrs("+Request.Form("MM_recordId")+",'"+Request.Form("AddressType")+"','"+UserName+"','"+ContactFirstName+"','"+ContactLastName+"','"+Request.Form("insUsr_ship_Fclty")+"',"+Session("insStaff_id")+",0,'E',0)}";
	rsShippingAddress.CursorType = 0;
	rsShippingAddress.CursorLocation = 2;
	rsShippingAddress.LockType = 3;	
	//Response.Redirect(rsShippingAddress.Source);
	rsShippingAddress.Open();
}

var rsShippingAddress = Server.CreateObject("ADODB.Recordset");
rsShippingAddress.ActiveConnection = MM_cnnASP02_STRING;
rsShippingAddress.Source = "{call dbo.cp_eqpsrv_ship_addrs("+Request.QueryString("intEquip_srv_id")+",'','','','',0,0,0,'Q',0)}";
rsShippingAddress.CursorType = 0;
rsShippingAddress.CursorLocation = 2;
rsShippingAddress.LockType = 3;
rsShippingAddress.Open();

var rsPhoneType = Server.CreateObject("ADODB.Recordset");
rsPhoneType.ActiveConnection = MM_cnnASP02_STRING;
rsPhoneType.Source = "{call dbo.cp_Phone_Type}";
rsPhoneType.CursorType = 0;
rsPhoneType.CursorLocation = 2;
rsPhoneType.LockType = 3;
rsPhoneType.Open();

if (rsShippingAddress.Fields.Item("insUser_Type_id").Value=="3") {
	var CheckHomeAddress = Server.CreateObject("ADODB.Command");
	CheckHomeAddress.ActiveConnection = MM_cnnASP02_STRING;
	CheckHomeAddress.CommandText = "dbo.cp_eqpsrv_homeaddrs_drv";
	CheckHomeAddress.CommandType = 4;
	CheckHomeAddress.CommandTimeout = 0;
	CheckHomeAddress.Prepared = true;
	CheckHomeAddress.Parameters.Append(CheckHomeAddress.CreateParameter("RETURN_VALUE", 3, 4));
	CheckHomeAddress.Parameters.Append(CheckHomeAddress.CreateParameter("@insUser_id", 3, 1,10000,rsShippingAddress.Fields.Item("insUser_id").Value));
	CheckHomeAddress.Parameters.Append(CheckHomeAddress.CreateParameter("@intRtnFlag", 2, 2));
	CheckHomeAddress.Execute();

	var rsHomeAddress = Server.CreateObject("ADODB.Recordset");
	rsHomeAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsHomeAddress.Source = "{call dbo.cp_eqpsrv_homeaddrs_drv("+rsShippingAddress.Fields.Item("insUser_id").Value+",0)}";
	rsHomeAddress.CursorType = 0;
	rsHomeAddress.CursorLocation = 2;
	rsHomeAddress.LockType = 3;
	rsHomeAddress.Open();
	
	var rsSchoolAddress = Server.CreateObject("ADODB.Recordset");
	rsSchoolAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsSchoolAddress.Source = "{call dbo.cp_eqpsrv_schaddrs_drv("+rsShippingAddress.Fields.Item("insUser_id").Value+",0)}";
	rsSchoolAddress.CursorType = 0;
	rsSchoolAddress.CursorLocation = 2;
	rsSchoolAddress.LockType = 3;
	rsSchoolAddress.Open();
	
	var rsWorkAddress = Server.CreateObject("ADODB.Recordset");
	rsWorkAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsWorkAddress.Source = "{call dbo.cp_eqpsrv_emplyadrs_drv("+rsShippingAddress.Fields.Item("insUser_id").Value+",0)}";
	rsWorkAddress.CursorType = 0;
	rsWorkAddress.CursorLocation = 2;
	rsWorkAddress.LockType = 3;
	rsWorkAddress.Open();
} else if (rsShippingAddress.Fields.Item("insUser_Type_id").Value=="4") {
	var rsInstitutionAddress = Server.CreateObject("ADODB.Recordset");
	rsInstitutionAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitutionAddress.Source = "{call dbo.cp_school_address("+rsShippingAddress.Fields.Item("insInstit_User_id").Value+",0,'','',0,'',0,'','','',0,'','','',0,'','','','','',1,'Q',0)}";
	rsInstitutionAddress.CursorType = 0;
	rsInstitutionAddress.CursorLocation = 2;
	rsInstitutionAddress.LockType = 3;
	rsInstitutionAddress.Open();
}
%>
<html>
<head>
	<title>Shipping Address</title>
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
				window.location.href='m009e0401.asp?intEquip_srv_id=<%=Request.QueryString("intEquip_srv_id")%>';
			break;
		}
	}
	</script>
	<script language="Javascript">
	var HomeArray = new Array(13);
	HomeArray[0] = "";
	HomeArray[1] = "";
	HomeArray[2] = "";
	HomeArray[3] = "";
	HomeArray[4] = "";
	HomeArray[5] = 0;
	HomeArray[6] = "";
	HomeArray[7] = "";
	HomeArray[8] = "";
	HomeArray[9] = 0;
	HomeArray[10] = "";
	HomeArray[11] = "";
	HomeArray[12] = "";
	
	var SchoolArray = new Array(13);
	SchoolArray[0] = "";
	SchoolArray[1] = "";
	SchoolArray[2] = "";
	SchoolArray[3] = "";
	SchoolArray[4] = "";
	SchoolArray[5] = 0;
	SchoolArray[6] = "";
	SchoolArray[7] = "";
	SchoolArray[8] = "";
	SchoolArray[9] = 0;
	SchoolArray[10] = "";
	SchoolArray[11] = "";
	SchoolArray[12] = "";
	
	var WorkArray = new Array(13);
	WorkArray[0] = "";
	WorkArray[1] = "";
	WorkArray[2] = "";
	WorkArray[3] = "";
	WorkArray[4] = "";
	WorkArray[5] = 0;
	WorkArray[6] = "";
	WorkArray[7] = "";
	WorkArray[8] = "";
	WorkArray[9] = 0;
	WorkArray[10] = "";
	WorkArray[11] = "";
	WorkArray[12] = "";
<%
if (rsShippingAddress.Fields.Item("insUser_Type_id").Value=="3") {
	if (!(CheckHomeAddress.Parameters.Item("@intRtnFlag").Value==null)) {
%>
		HomeArray[0] = "<%=Trim(rsHomeAddress.Fields.Item("chvAddress").Value.replace(/\r\n|\r|\n/g, ' '))%>";
		HomeArray[1] = "<%=Trim(rsHomeAddress.Fields.Item("chvCity").Value)%>";
		HomeArray[2] = "<%=Trim(rsHomeAddress.Fields.Item("chvProv").Value)%>";
		HomeArray[3] = "<%=Trim(rsHomeAddress.Fields.Item("chvCountry").Value)%>";
		HomeArray[4] = "<%=FormatPostalCode(rsHomeAddress.Fields.Item("chvPostal_zip").Value)%>";
		HomeArray[5] = "<%=Trim(rsHomeAddress.Fields.Item("intPhone_Type_1").Value)%>";
		HomeArray[6] = "<%=Trim(rsHomeAddress.Fields.Item("chvPhone1_Arcd").Value)%>";
		HomeArray[7] = "<%=FormatPhoneNumberOnly(rsHomeAddress.Fields.Item("chvPhone1_Num").Value)%>";
		HomeArray[8] = "<%=Trim(rsHomeAddress.Fields.Item("chvPhone1_Ext").Value)%>";
		HomeArray[9] = "<%=Trim(rsHomeAddress.Fields.Item("intPhone_Type_2").Value)%>";
		HomeArray[10] = "<%=Trim(rsHomeAddress.Fields.Item("chvPhone2_Arcd").Value)%>";
		HomeArray[11] = "<%=FormatPhoneNumberOnly(rsHomeAddress.Fields.Item("chvPhone2_Num").Value)%>";
		HomeArray[12] = "<%=Trim(rsHomeAddress.Fields.Item("chvPhone2_Ext").Value)%>";	
<%	
	}
	if (!rsSchoolAddress.EOF) {
%>
		SchoolArray[0] = "<%=Trim(rsSchoolAddress.Fields.Item("chvAddress").Value.replace(/\r\n|\r|\n/g, ' '))%>";
		SchoolArray[1] = "<%=Trim(rsSchoolAddress.Fields.Item("chvCity").Value)%>";
		SchoolArray[2] = "<%=Trim(rsSchoolAddress.Fields.Item("chvProv").Value)%>";
		SchoolArray[3] = "<%=Trim(rsSchoolAddress.Fields.Item("chvCountry").Value)%>";
		SchoolArray[4] = "<%=FormatPostalCode(rsSchoolAddress.Fields.Item("chvPostal_zip").Value)%>";
		SchoolArray[5] = "<%=Trim(rsSchoolAddress.Fields.Item("intPhone_Type_1").Value)%>";
		SchoolArray[6] = "<%=Trim(rsSchoolAddress.Fields.Item("chvPhone1_Arcd").Value)%>";
		SchoolArray[7] = "<%=FormatPhoneNumberOnly(rsSchoolAddress.Fields.Item("chvPhone1_Num").Value)%>";
		SchoolArray[8] = "<%=Trim(rsSchoolAddress.Fields.Item("chvPhone1_Ext").Value)%>";
		SchoolArray[9] = "<%=Trim(rsSchoolAddress.Fields.Item("intPhone_Type_2").Value)%>";
		SchoolArray[10] = "<%=Trim(rsSchoolAddress.Fields.Item("chvPhone2_Arcd").Value)%>";
		SchoolArray[11] = "<%=FormatPhoneNumberOnly(rsSchoolAddress.Fields.Item("chvPhone2_Num").Value)%>";
		SchoolArray[12] = "<%=Trim(rsSchoolAddress.Fields.Item("chvPhone2_Ext").Value)%>";		
<%
	}	
	if (!rsWorkAddress.EOF) {
%>
		WorkArray[0] = "<%=Trim(String(rsWorkAddress.Fields.Item("chvAddress").Value).replace(/\r\n|\r|\n/g, ' '))%>";
		WorkArray[1] = "<%=Trim(rsWorkAddress.Fields.Item("chvCity").Value)%>";
		WorkArray[2] = "<%=Trim(rsWorkAddress.Fields.Item("chvProv").Value)%>";
		WorkArray[3] = "<%=Trim(rsWorkAddress.Fields.Item("chvCountry").Value)%>";
		WorkArray[4] = "<%=FormatPostalCode(rsWorkAddress.Fields.Item("chvPostal_zip").Value)%>";
		WorkArray[5] = "<%=Trim(rsWorkAddress.Fields.Item("intPhone_Type_1").Value)%>";
		WorkArray[6] = "<%=Trim(rsWorkAddress.Fields.Item("chvPhone1_Arcd").Value)%>";
		WorkArray[7] = "<%=FormatPhoneNumberOnly(rsWorkAddress.Fields.Item("chvPhone1_Num").Value)%>";
		WorkArray[8] = "<%=Trim(rsWorkAddress.Fields.Item("chvPhone1_Ext").Value)%>";
		WorkArray[9] = "<%=Trim(rsWorkAddress.Fields.Item("intPhone_Type_2").Value)%>";
		WorkArray[10] = "<%=Trim(rsWorkAddress.Fields.Item("chvPhone2_Arcd").Value)%>";
		WorkArray[11] = "<%=FormatPhoneNumberOnly(rsWorkAddress.Fields.Item("chvPhone2_Num").Value)%>";
		WorkArray[12] = "<%=Trim(rsWorkAddress.Fields.Item("chvPhone2_Ext").Value)%>";		
<%
	}
} else if (rsShippingAddress.Fields.Item("insUser_Type_id").Value=="4") {
	if (!rsInstitutionAddress.EOF) {
%>
		SchoolArray[0] = "<%=Trim(String(rsInstitutionAddress.Fields.Item("chvAddress").Value).replace(/\r\n|\r|\n/g, ' '))%>";
		SchoolArray[1] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvCity").Value)%>";
		SchoolArray[2] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvProvince").Value)%>";
		SchoolArray[3] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvcntry_name").Value)%>";
		SchoolArray[4] = "<%=FormatPostalCode(rsInstitutionAddress.Fields.Item("chvPostal_zip").Value)%>";
		SchoolArray[5] = "<%=Trim(rsInstitutionAddress.Fields.Item("intPhone_Type_1").Value)%>";
		SchoolArray[6] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvPhone1_Arcd").Value)%>";
		SchoolArray[7] = "<%=FormatPhoneNumberOnly(rsInstitutionAddress.Fields.Item("chvPhone1_Num").Value)%>";
		SchoolArray[8] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvPhone1_Ext").Value)%>";
		SchoolArray[9] = "<%=Trim(rsInstitutionAddress.Fields.Item("intPhone_Type_2").Value)%>";
		SchoolArray[10] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvPhone2_Arcd").Value)%>";
		SchoolArray[11] = "<%=FormatPhoneNumberOnly(rsInstitutionAddress.Fields.Item("chvPhone2_Num").Value)%>";
		SchoolArray[12] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvPhone2_Ext").Value)%>";		
<%
	}
}
%>	
	function ChangeAddressType(){
		switch (document.frm0402.AddressType.value) {
			//Home
			case "1":
				document.frm0402.StreetAddress.value=HomeArray[0];
				document.frm0402.City.value=HomeArray[1];
				document.frm0402.ProvinceState.value=HomeArray[2];								
				document.frm0402.Country.value=HomeArray[3];
				document.frm0402.PostalCode.value=HomeArray[4];								
				document.frm0402.PrimaryPhoneType.value=HomeArray[5];
				document.frm0402.PrimaryPhoneAreaCode.value=HomeArray[6];
				document.frm0402.PrimaryPhoneNumber.value=HomeArray[7];
				document.frm0402.PrimaryPhoneExtension.value=HomeArray[8];
				document.frm0402.SecondaryPhoneType.value=HomeArray[9];
				document.frm0402.SecondaryPhoneAreaCode.value=HomeArray[10];
				document.frm0402.SecondaryPhoneNumber.value=HomeArray[11];
				document.frm0402.SecondaryPhoneExtension.value=HomeArray[12];
			break;
			//PostSec
			case "2":
				document.frm0402.StreetAddress.value=SchoolArray[0];
				document.frm0402.City.value=SchoolArray[1];
				document.frm0402.ProvinceState.value=SchoolArray[2];		
				document.frm0402.Country.value=SchoolArray[3];
				document.frm0402.PostalCode.value=SchoolArray[4];								
				document.frm0402.PrimaryPhoneType.value=SchoolArray[5];
				document.frm0402.PrimaryPhoneAreaCode.value=SchoolArray[6];
				document.frm0402.PrimaryPhoneNumber.value=SchoolArray[7];
				document.frm0402.PrimaryPhoneExtension.value=SchoolArray[8];
				document.frm0402.SecondaryPhoneType.value=SchoolArray[9];
				document.frm0402.SecondaryPhoneAreaCode.value=SchoolArray[10];
				document.frm0402.SecondaryPhoneNumber.value=SchoolArray[11];
				document.frm0402.SecondaryPhoneExtension.value=SchoolArray[12];
			break;
			//Work
			case "3":
				document.frm0402.StreetAddress.value=WorkArray[0];
				document.frm0402.City.value=WorkArray[1];
				document.frm0402.ProvinceState.value=WorkArray[2];	
				document.frm0402.Country.value=WorkArray[3];
				document.frm0402.PostalCode.value=WorkArray[4];								
				document.frm0402.PrimaryPhoneType.value=WorkArray[5];
				document.frm0402.PrimaryPhoneAreaCode.value=WorkArray[6];
				document.frm0402.PrimaryPhoneNumber.value=WorkArray[7];
				document.frm0402.PrimaryPhoneExtension.value=WorkArray[8];
				document.frm0402.SecondaryPhoneType.value=WorkArray[9];
				document.frm0402.SecondaryPhoneAreaCode.value=WorkArray[10];
				document.frm0402.SecondaryPhoneNumber.value=WorkArray[11];
				document.frm0402.SecondaryPhoneExtension.value=WorkArray[12];
			break;
			//Other
			case "0":
				document.frm0402.StreetAddress.value="";
				document.frm0402.City.value="";
				document.frm0402.ProvinceState.value="";								
				document.frm0402.Country.value="";
				document.frm0402.PostalCode.value="";								
				document.frm0402.PrimaryPhoneType.value="";
				document.frm0402.PrimaryPhoneAreaCode.value="";
				document.frm0402.PrimaryPhoneNumber.value="";
				document.frm0402.PrimaryPhoneExtension.value="";
				document.frm0402.SecondaryPhoneType.value="";
				document.frm0402.SecondaryPhoneAreaCode.value="";
				document.frm0402.SecondaryPhoneNumber.value="";
				document.frm0402.SecondaryPhoneExtension.value="";
			break;	
		}	
	}
	
	function Init(){
		document.frm0402.AddressType.focus();
		//ChangeAddressType();
	}

	function PrintShippingLabel(){
		document.frm0402.action = "m009e0404.asp?intEquip_srv_id=<%=Request.QueryString("intEquip_srv_id")%>";
		document.frm0402.target = "_blank";
		document.frm0402.submit();	
	}
			
	function Save(){
		document.frm0402.action = "<%=MM_editAction%>";
		document.frm0402.target = "_self";
		document.frm0402.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0402">
<h5>Shipping Address</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Address Type:</td>
		<td nowrap><select name="AddressType" tabindex="1" onChange="ChangeAddressType();" accesskey="F">
			<option value="1" <%=((rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value=="1")?"SELECTED":"")%>>Home
			<option value="2" <%=((rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value=="2")?"SELECTED":"")%>>Post-Sec Institution
			<option value="3" <%=((rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value=="3")?"SELECTED":"")%>>Work
			<option value="0" <%=((((rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value>"3") || (rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value<"1")))?"SELECTED":"")%>>Other
		</select></td>
		<td nowrap>Facility:</td>
		<td nowrap><input type="text" name="Facility" value="<%=(rsShippingAddress.Fields.Item("chvUsr_ship_Fclty").Value)%>" readonly size="30" tabindex="2"></td>		
	</tr>
	<tr>
		<td nowrap>User Name:</td>
		<td nowrap colspan="3"><input type="text" name="UserName" value="<%=(rsShippingAddress.Fields.Item("chvUsr_Name").Value)%>" tabindex="3" size="40"></td>
	</tr>
	<tr>
		<td nowrap>Contact First Name:</td>
		<td nowrap><input type="text" name="ContactFirstName" value="<%=(rsShippingAddress.Fields.Item("chvCo_Fstname").Value)%>" tabindex="4"></td>
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="ContactLastName" value="<%=(rsShippingAddress.Fields.Item("chvCo_Lstname").Value)%>" tabindex="5"></td>
	</tr>
	<tr>
		<td nowrap valign="top">Street Address:</td>
		<td nowrap valign="top"><textarea name="StreetAddress" cols="30" rows="3" tabindex="6" readonly></textarea></td>
		<td colspan="2"></td>
	</tr>
	<tr> 
		<td nowrap>City:</td>
		<td nowrap><input type="text" name="City" tabindex="7" readonly maxlength="50"></td>
		<td nowrap>Province/State:</td>
		<td nowrap><input type="text" name="ProvinceState" tabindex="8" readonly size="2"></td>
    </tr>	
	<tr>
		<td nowrap>Country:</td>
		<td nowrap><input type="text" name="Country" readonly tabindex="9"></td>
		<td nowrap>Postal Code:</td>
		<td nowrap><input type="text" name="PostalCode" readonly tabindex="10" size="10" maxlength="7"></td>
    </tr>
    <tr> 
		<td nowrap>Primary Phone:</td>
		<td nowrap colspan="3"> 
			<select name="PrimaryPhoneType" tabindex="11">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>"><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			rsPhoneType.MoveFirst();
			%>
			</select>
			<input type="text" name="PrimaryPhoneAreaCode" tabindex="12" readonly size="4">
			<input type="text" name="PrimaryPhoneNumber" tabindex="13" readonly size="9" maxlength="8">Ext
			<input type="text" name="PrimaryPhoneExtension" tabindex="14" readonly size="4">
		</td>
    </tr>
    <tr> 
		<td nowrap>Secondary Phone:</td>
		<td nowrap colspan="3">
			<select name="SecondaryPhoneType" tabindex="15">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>"><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			%>
			</select>
			<input type="text" name="SecondaryPhoneAreaCode" tabindex="16" readonly size="4">
			<input type="text" name="SecondaryPhoneNumber" tabindex="17" readonly size="9" maxlength="8">Ext
			<input type="text" name="SecondaryPhoneExtension" tabindex="18" readonly size="4">
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="19" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Print Shipping Label" tabindex="20" onClick="PrintShippingLabel();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="21" onClick="window.location.href='m009e0401.asp?intEquip_srv_id=<%=Request.QueryString("intEquip_srv_id")%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="update">
<input type="hidden" name="MM_recordId" value="<%=Request.QueryString("intEquip_srv_id")%>">
<input type="hidden" name="insUsr_ship_Fclty" value="<%=(rsShippingAddress.Fields.Item("insUsr_ship_Fclty").Value)%>">
</form>
</body>
</html>
<%
rsShippingAddress.Close();
%>
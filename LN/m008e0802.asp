<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsLoanDetail = Server.CreateObject("ADODB.Recordset");
rsLoanDetail.ActiveConnection = MM_cnnASP02_STRING;
rsLoanDetail.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoanDetail.CursorType = 0;
rsLoanDetail.CursorLocation = 2;
rsLoanDetail.LockType = 3;
rsLoanDetail.Open();

var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_get_loan_ship_name("+ Request.QueryString("intLoan_Req_id") + ",0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var intBOShip_dtl_id = 0;
if (!rsLoan.EOF) {
	if (rsLoanDetail.Fields.Item("intBOShip_dtl_id").Value != null) intBOShip_dtl_id = Number(rsLoanDetail.Fields.Item("intBOShip_dtl_id").Value);
}			

if (String(Request("MM_action")) == "update") {
	var UserFirstName = String(Request.Form("UserFirstName")).replace(/'/g, "''");
	var UserLastName = String(Request.Form("UserLastName")).replace(/'/g, "''");
	var ContactFirstName = String(Request.Form("ContactFirstName")).replace(/'/g, "''");
	var ContactLastName = String(Request.Form("ContactLastName")).replace(/'/g, "''");
	var StreetAddress = String(Request.Form("StreetAddress")).replace(/'/g, "''");				
	var City = String(Request.Form("City")).replace(/'/g, "''");
	var rsUpdateAddress = Server.CreateObject("ADODB.Recordset");
	rsUpdateAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsUpdateAddress.Source = "{call dbo.cp_loan_ship_address("+Request.Form("MM_recordId")+",0,'"+Request.Form("AddressType")+"','"+UserFirstName+"','"+UserLastName+"','"+ContactFirstName+"','"+ContactLastName+"',"+Request.Form("AddressID")+",'"+StreetAddress+"','"+Trim(City)+"',"+Request.Form("ProvinceState")+",'"+Trim(Request.Form("PostalCode"))+"',"+Request.Form("PrimaryPhoneType")+",'"+Trim(Request.Form("PrimaryPhoneAreaCode"))+"','"+Trim(Request.Form("PrimaryPhoneNumber"))+"','"+Trim(Request.Form("PrimaryPhoneExtension"))+"',"+Request.Form("SecondaryPhoneType")+",'"+Request.Form("SecondaryPhoneAreaCode")+"','"+Trim(Request.Form("SecondaryPhoneNumber"))+"','"+Trim(Request.Form("SecondaryPhoneExtension"))+"',0,'','','','"+Request.Form("Email")+"','',0,'E',0)}";
	rsUpdateAddress.CursorType = 0;
	rsUpdateAddress.CursorLocation = 2;
	rsUpdateAddress.LockType = 3;	
	rsUpdateAddress.Open();
}

if (String(Request("MM_action")) == "insert") {
	var UserFirstName = String(Request.Form("UserFirstName")).replace(/'/g, "''");
	var UserLastName = String(Request.Form("UserLastName")).replace(/'/g, "''");
	var ContactFirstName = String(Request.Form("ContactFirstName")).replace(/'/g, "''");
	var ContactLastName = String(Request.Form("ContactLastName")).replace(/'/g, "''");
	var StreetAddress = String(Request.Form("StreetAddress")).replace(/'/g, "''");				
	var City = String(Request.Form("City")).replace(/'/g, "''");
	var rsInsertAddress = Server.CreateObject("ADODB.Recordset");
	rsInsertAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsInsertAddress.Source = "{call dbo.cp_loan_ship_address("+Request.Form("MM_recordId")+",0,'"+Request.Form("AddressType")+"','"+UserFirstName+"','"+UserLastName+"','"+ContactFirstName+"','"+ContactLastName+"',0,'"+StreetAddress+"','"+Trim(City)+"',"+Request.Form("ProvinceState")+",'"+Trim(Request.Form("PostalCode"))+"',"+Request.Form("PrimaryPhoneType")+",'"+Trim(Request.Form("PrimaryPhoneAreaCode"))+"','"+Trim(Request.Form("PrimaryPhoneNumber"))+"','"+Trim(Request.Form("PrimaryPhoneExtension"))+"',"+Request.Form("SecondaryPhoneType")+",'"+Request.Form("SecondaryPhoneAreaCode")+"','"+Trim(Request.Form("SecondaryPhoneNumber"))+"','"+Trim(Request.Form("SecondaryPhoneExtension"))+"',0,'','','','"+Request.Form("Email")+"','',0,'A',0)}";
	rsInsertAddress.CursorType = 0;
	rsInsertAddress.CursorLocation = 2;
	rsInsertAddress.LockType = 3;
//	Response.Redirect(rsInsertAddress.Source);
	rsInsertAddress.Open();
}

var rsShippingAddress = Server.CreateObject("ADODB.Recordset");
rsShippingAddress.ActiveConnection = MM_cnnASP02_STRING;
rsShippingAddress.Source = "{call dbo.cp_loan_ship_address("+intBOShip_dtl_id+",0,'','','','','',0,'','',0,'',0,'','','',0,'','','',0,'','','','','',0,'Q',0)}";
rsShippingAddress.CursorType = 0;
rsShippingAddress.CursorLocation = 2;
rsShippingAddress.LockType = 3;
rsShippingAddress.Open();

//var IsNew = ((rsShippingAddress.EOF)?true:false);
if (rsShippingAddress.EOF) {
	IsNew = true;
} else {
	if ((rsShippingAddress.Fields.Item("intaddress_id").Value==null) || (rsShippingAddress.Fields.Item("intaddress_id").Value==0)) {
		IsNew = true;		
	} else {
		IsNew = false;		
	}
}

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

if (rsLoan.Fields.Item("insEq_user_type").Value=="3") {
	var CheckHomeAddress = Server.CreateObject("ADODB.Command");
	CheckHomeAddress.ActiveConnection = MM_cnnASP02_STRING;
	CheckHomeAddress.CommandText = "dbo.cp_eqpsrv_homeaddrs_drv";
	CheckHomeAddress.CommandType = 4;
	CheckHomeAddress.CommandTimeout = 0;
	CheckHomeAddress.Prepared = true;
	CheckHomeAddress.Parameters.Append(CheckHomeAddress.CreateParameter("RETURN_VALUE", 3, 4));
	CheckHomeAddress.Parameters.Append(CheckHomeAddress.CreateParameter("@insUser_id", 3, 1,10000,rsLoan.Fields.Item("intEq_user_id").Value));
	CheckHomeAddress.Parameters.Append(CheckHomeAddress.CreateParameter("@intRtnFlag", 2, 2));
	CheckHomeAddress.Execute();
		
	var rsHomeAddress = Server.CreateObject("ADODB.Recordset");
	rsHomeAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsHomeAddress.Source = "{call dbo.cp_eqpsrv_homeaddrs_drv("+rsLoan.Fields.Item("intEq_user_id").Value+",0)}";
	rsHomeAddress.CursorType = 0;
	rsHomeAddress.CursorLocation = 2;
	rsHomeAddress.LockType = 3;
	rsHomeAddress.Open();
		
	var rsSchoolAddress = Server.CreateObject("ADODB.Recordset");
	rsSchoolAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsSchoolAddress.Source = "{call dbo.cp_eqpsrv_schaddrs_drv("+rsLoan.Fields.Item("intEq_user_id").Value+",0)}";
	rsSchoolAddress.CursorType = 0;
	rsSchoolAddress.CursorLocation = 2;
	rsSchoolAddress.LockType = 3;
	rsSchoolAddress.Open();
	
	var rsWorkAddress = Server.CreateObject("ADODB.Recordset");
	rsWorkAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsWorkAddress.Source = "{call dbo.cp_eqpsrv_emplyadrs_drv("+rsLoan.Fields.Item("intEq_user_id").Value+",0)}";
	rsWorkAddress.CursorType = 0;
	rsWorkAddress.CursorLocation = 2;
	rsWorkAddress.LockType = 3;
	rsWorkAddress.Open();
} else if (rsLoan.Fields.Item("insEq_user_type").Value=="4") {
	var rsInstitutionAddress = Server.CreateObject("ADODB.Recordset");
	rsInstitutionAddress.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitutionAddress.Source = "{call dbo.cp_school_address("+rsLoan.Fields.Item("intEq_user_id").Value+",0,'','',0,'',0,'','','',0,'','','',0,'','','','','',1,'Q',0)}";
	rsInstitutionAddress.CursorType = 0;
	rsInstitutionAddress.CursorLocation = 2;
	rsInstitutionAddress.LockType = 3;
	rsInstitutionAddress.Open();
}
%>
<html>
<head>
	<title>Backorder Shipping Address</title>
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
				window.location.href='m008e0801.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>';
			break;
		}
	}
	</script>	
	<script language="Javascript">
	var HomeArray = new Array(13);
	HomeArray[0] = "";
	HomeArray[1] = "";
	HomeArray[2] = 101;
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
	SchoolArray[2] = 101;
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
	WorkArray[2] = 101;
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
//if (!IsNew) {
	if (rsLoan.Fields.Item("insEq_user_type").Value=="3") {
		if (!(CheckHomeAddress.Parameters.Item("@intRtnFlag").Value==null)) {
%>
			HomeArray[0] = "<%=Trim(rsHomeAddress.Fields.Item("chvAddress").Value.replace(/\r\n|\r|\n/g, ' '))%>";
			HomeArray[1] = "<%=Trim(rsHomeAddress.Fields.Item("chvCity").Value)%>";
			HomeArray[2] = "<%=(((rsHomeAddress.Fields.Item("insProv_State_id").Value==0)||(rsHomeAddress.Fields.Item("insProv_State_id").Value=="")||(rsHomeAddress.Fields.Item("insProv_State_id").Value==null))?"101":rsHomeAddress.Fields.Item("insProv_State_id").Value)%>";
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
			SchoolArray[2] = "<%=(((rsSchoolAddress.Fields.Item("insProv_State_id").Value==0)||(rsSchoolAddress.Fields.Item("insProv_State_id").Value=="")||(rsSchoolAddress.Fields.Item("insProv_State_id").Value==null))?"101":rsSchoolAddress.Fields.Item("insProv_State_id").Value)%>";			
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
			WorkArray[0] = "<%=Trim(rsWorkAddress.Fields.Item("chvAddress").Value.replace(/\r\n|\r|\n/g, ' '))%>";
			WorkArray[1] = "<%=Trim(rsWorkAddress.Fields.Item("chvCity").Value)%>";
			WorkArray[2] = "<%=(((rsWorkAddress.Fields.Item("insProv_State_id").Value==0)||(rsWorkAddress.Fields.Item("insProv_State_id").Value=="")||(rsWorkAddress.Fields.Item("insProv_State_id").Value==null))?"101":rsWorkAddress.Fields.Item("insProv_State_id").Value)%>";												
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
	} else if (rsLoan.Fields.Item("insEq_user_type").Value=="4") {
		if (!rsInstitutionAddress.EOF) {
%>
			SchoolArray[0] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvAddress").Value.replace(/\r\n|\r|\n/g, ' '))%>";
			SchoolArray[1] = "<%=Trim(rsInstitutionAddress.Fields.Item("chvCity").Value)%>";
			SchoolArray[2] = "<%=(((rsInstitutionAddress.Fields.Item("intprvst_id").Value==0)||(rsInstitutionAddress.Fields.Item("intprvst_id").Value=="")||(rsInstitutionAddress.Fields.Item("intprvst_id").Value==null))?"101":rsInstitutionAddress.Fields.Item("intprvst_id").Value)%>";
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
//}
%>	
	function ChangeAddressType(){
		switch (document.frm0802.AddressType.value) {
			//Home
			case "1":
				document.frm0802.StreetAddress.value=HomeArray[0];
				document.frm0802.City.value=HomeArray[1];
				document.frm0802.ProvinceState.value=HomeArray[2];								
				document.frm0802.Country.value=HomeArray[3];
				document.frm0802.PostalCode.value=HomeArray[4];								
				document.frm0802.PrimaryPhoneType.value=HomeArray[5];
				document.frm0802.PrimaryPhoneAreaCode.value=HomeArray[6];
				document.frm0802.PrimaryPhoneNumber.value=HomeArray[7];
				document.frm0802.PrimaryPhoneExtension.value=HomeArray[8];
				document.frm0802.SecondaryPhoneType.value=HomeArray[9];
				document.frm0802.SecondaryPhoneAreaCode.value=HomeArray[10];
				document.frm0802.SecondaryPhoneNumber.value=HomeArray[11];
				document.frm0802.SecondaryPhoneExtension.value=HomeArray[12];
			break;
			//PostSec
			case "2":
				document.frm0802.StreetAddress.value=SchoolArray[0];
				document.frm0802.City.value=SchoolArray[1];
				document.frm0802.ProvinceState.value=SchoolArray[2];		
				document.frm0802.Country.value=SchoolArray[3];
				document.frm0802.PostalCode.value=SchoolArray[4];								
				document.frm0802.PrimaryPhoneType.value=SchoolArray[5];
				document.frm0802.PrimaryPhoneAreaCode.value=SchoolArray[6];
				document.frm0802.PrimaryPhoneNumber.value=SchoolArray[7];
				document.frm0802.PrimaryPhoneExtension.value=SchoolArray[8];
				document.frm0802.SecondaryPhoneType.value=SchoolArray[9];
				document.frm0802.SecondaryPhoneAreaCode.value=SchoolArray[10];
				document.frm0802.SecondaryPhoneNumber.value=SchoolArray[11];
				document.frm0802.SecondaryPhoneExtension.value=SchoolArray[12];
			break;
			//Work
			case "3":
				document.frm0802.StreetAddress.value=WorkArray[0];
				document.frm0802.City.value=WorkArray[1];
				document.frm0802.ProvinceState.value=WorkArray[2];	
				document.frm0802.Country.value=WorkArray[3];
				document.frm0802.PostalCode.value=WorkArray[4];								
				document.frm0802.PrimaryPhoneType.value=WorkArray[5];
				document.frm0802.PrimaryPhoneAreaCode.value=WorkArray[6];
				document.frm0802.PrimaryPhoneNumber.value=WorkArray[7];
				document.frm0802.PrimaryPhoneExtension.value=WorkArray[8];
				document.frm0802.SecondaryPhoneType.value=WorkArray[9];
				document.frm0802.SecondaryPhoneAreaCode.value=WorkArray[10];
				document.frm0802.SecondaryPhoneNumber.value=WorkArray[11];
				document.frm0802.SecondaryPhoneExtension.value=WorkArray[12];
			break;
			//Other
			case "0":
				document.frm0802.StreetAddress.value="";
				document.frm0802.City.value="";
				document.frm0802.ProvinceState.value=101;								
				document.frm0802.Country.value="";
				document.frm0802.PostalCode.value="";								
				document.frm0802.PrimaryPhoneType.value=0;
				document.frm0802.PrimaryPhoneAreaCode.value="";
				document.frm0802.PrimaryPhoneNumber.value="";
				document.frm0802.PrimaryPhoneExtension.value="";
				document.frm0802.SecondaryPhoneType.value=0;
				document.frm0802.SecondaryPhoneAreaCode.value="";
				document.frm0802.SecondaryPhoneNumber.value="";
				document.frm0802.SecondaryPhoneExtension.value="";
			break;	
		}	
	}
	function Init(){
	<%
	if (!rsShippingAddress.EOF) {
	%>
		document.frm0802.AddressType.focus();
	<%
		if (IsNew) {
	%>
		ChangeAddressType();	
	<%
		}
	}
	%>
	}
	
	function Save(){
		if (!CheckPostalCode(document.frm0802.PostalCode.value)){
			alert("Invalid Postal Code.");
			document.frm0802.PostalCode.focus();
			return ;
		}
		var tempPC = document.frm0802.PostalCode.value;
		tempPC = tempPC.toUpperCase();
		document.frm0802.PostalCode.value = tempPC;						
		document.frm0802.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0802">
<h5>Backorder Shipping Address</h5>
<hr>
<%
if (rsShippingAddress.EOF) {
%>
<i>Please go to Method page and save first, before entering shipping address.</i>
<%
} else {
%>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Address Type</td>
		<td nowrap colspan="3"><select name="AddressType" tabindex="1" onChange="ChangeAddressType();" accesskey="F">
			<option value="1" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value=="1")?"SELECTED":""))%>>Home
			<option value="2" <%if (!IsNew) {Response.Write(((rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value=="2")?"SELECTED":""))} else {Response.Write("SELECTED")}%>>Post-Sec Institution			
			<option value="3" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value=="3")?"SELECTED":""))%>>Work
			<option value="0" <%if (!IsNew) Response.Write((((rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value>"3") || (rsShippingAddress.Fields.Item("chrShip_Adrs_type").Value<"1"))?"SELECTED":""))%>>Other
		</select></td>
	</tr>
	<tr>
		<td nowrap>User First Name:</td>
		<td nowrap><input type="text" name="UserFirstName" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvUsr_Fstname").Value:"")%>" tabindex="3"></td>
		<td nowrap>User Last Name:</td>
		<td nowrap><input type="text" name="UserLastName" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvUsr_Lstname").Value:"")%>" tabindex="4"></td>
	</tr>
	<tr>
		<td nowrap>Contact First Name:</td>
		<td nowrap><input type="text" name="ContactFirstName" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvCo_Fstname").Value:"")%>" tabindex="5"></td>
		<td nowrap>Contact Last Name:</td>
		<td nowrap><input type="text" name="ContactLastName" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvCo_Lstname").Value:"")%>" tabindex="6"></td>
	</tr>
	<tr> 
		<td nowrap valign="top">Street Address:</td>
		<td nowrap colspan="3"><textarea name="StreetAddress" cols="30" rows="3" tabindex="7"><%=((!IsNew)?rsShippingAddress.Fields.Item("chvAddress").Value:"")%></textarea></td>
	</tr>
	<tr> 
		<td nowrap>City:</td>
		<td nowrap><input type="text" name="City" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvCity").Value:"")%>" maxlength="50" tabindex="8"></td>
		<td nowrap>Province/State:</td>
		<td nowrap>
			<select name="ProvinceState" tabindex="9">
			<% 
			while (!rsProvince.EOF) {
			%>
				<option value="<%=(rsProvince.Fields.Item("intprvst_id").Value)%>" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("intprvst_id").Value==rsProvince.Fields.Item("intprvst_id").Value)?"SELECTED":""))%>><%=(rsProvince.Fields.Item("chrprvst_abbv").Value)%>
			<%
				rsProvince.MoveNext();
			}
			%>
        </select></td>
    </tr>	
	<tr>
		<td nowrap>Country:</td>
		<td nowrap><input type="text" name="Country" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvcntry_name").Value:"")%>" tabindex="10" readonly></td>
		<td nowrap>Postal Code:</td>
		<td nowrap><input type="text" name="PostalCode" value="<%=((!IsNew)?FormatPostalCode(rsShippingAddress.Fields.Item("chvPostal_zip").Value):"")%>" tabindex="11" size="10" maxlength="7" onChange="FormatPostalCode(this);"></td>
    </tr>
    <tr> 
		<td nowrap>Primary Phone:</td>
		<td nowrap colspan="3"> 
			<select name="PrimaryPhoneType" tabindex="12">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("intPhone_Type_1").Value==rsPhoneType.Fields.Item("intPhone_type_id").Value)?"SELECTED":""))%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			rsPhoneType.MoveFirst();
			%>
			</select>
			<select name="PrimaryPhoneAreaCode" tabindex="13">
				<option value="" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("chvPhone1_Arcd").Value=="")?"SELECTED":""))%>>			
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("chvPhone1_Arcd").Value==rsAreaCode.Fields.Item("chvAC_num").Value)?"SELECTED":""))%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			rsAreaCode.MoveFirst();
			%>			
			</select>
			<input type="text" name="PrimaryPhoneNumber" value="<%=((!IsNew)?FormatPhoneNumberOnly(rsShippingAddress.Fields.Item("chvPhone1_Num").Value):"")%>" size="9" tabindex="14" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)" >
			Ext <input type="text" name="PrimaryPhoneExtension" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvPhone1_Ext").Value:"")%>" size="4" tabindex="15" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>Secondary Phone:</td>
		<td nowrap colspan="3">
			<select name="SecondaryPhoneType" tabindex="16">
				<% 
				while (!rsPhoneType.EOF) {
				%>
					<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("intPhone_Type_2").Value==rsPhoneType.Fields.Item("intPhone_type_id").Value)?"SELECTED":""))%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
				<%
					rsPhoneType.MoveNext();
				}
				%>
			</select>
			<select name="SecondaryPhoneAreaCode" tabindex="17">
				<option value="" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("chvPhone2_Arcd").Value=="")?"SELECTED":""))%>>
			<%
			while (!rsAreaCode.EOF) {
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%if (!IsNew) Response.Write(((rsShippingAddress.Fields.Item("chvPhone2_Arcd").Value==rsAreaCode.Fields.Item("chvAC_num").Value)?"SELECTED":""))%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>			
			</select>
			<input type="text" name="SecondaryPhoneNumber" value="<%=((!IsNew)?FormatPhoneNumberOnly(rsShippingAddress.Fields.Item("chvPhone2_Num").Value):"")%>" size="9" tabindex="18" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">
			Ext <input type="text" name="SecondaryPhoneExtension" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvPhone2_Ext").Value:"")%>" size="4" tabindex="19" onKeypress="AllowNumericOnly();">
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="20" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="21" onClick="window.location.href='m008e0801.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="AddressID" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("intaddress_id").Value:"0")%>">
<input type="hidden" name="Notes">
<input type="hidden" name="EMail" value="<%=((!IsNew)?rsShippingAddress.Fields.Item("chvEmail").Value:"")%>">
<input type="hidden" name="MM_action" value="<%=((IsNew)?"insert":"update")%>">
<input type="hidden" name="MM_recordId" value="<%=intBOShip_dtl_id%>">
<%
}
%>
</form>
</body>
</html>
<%
rsLoan.Close();
rsShippingAddress.Close();
rsProvince.Close();
rsPhoneType.Close();
rsAreaCode.Close();
%>
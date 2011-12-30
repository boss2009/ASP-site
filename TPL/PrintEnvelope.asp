<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
Response.ContentType = "application/msword"
if (Request.QueryString("RecipientType")=="Contact") {
	var rsContact = Server.CreateObject("ADODB.Recordset");
	rsContact.ActiveConnection = MM_cnnASP02_STRING;
	rsContact.Source = "{call dbo.cp_contacts("+Request.QueryString("To")+",0,'','','',0,0,0,1,0,'',1,'Q',0)}"
	rsContact.CursorType = 0;
	rsContact.CursorLocation = 2;
	rsContact.LockType = 3;
	rsContact.Open();
	if (!rsContact.EOF) {
		var To_Whole_Name = rsContact("chvFst_Name") + " " + rsContact("chvLst_Name") + "<br>";
		switch (String(rsContact.Fields.Item("intWork_type_id").Value)) {
			case "12":
				var rsInstitutionAddress = Server.CreateObject("ADODB.Recordset");
				rsInstitutionAddress.ActiveConnection = MM_cnnASP02_STRING;
				rsInstitutionAddress.Source = "{call dbo.cp_school_address("+ rsContact.Fields.Item("insWork_id").Value + ",0,'','',0,'',0,'','','',0,'','','',0,'','','','','',1,'Q',0)}";
				rsInstitutionAddress.CursorType = 0;
				rsInstitutionAddress.CursorLocation = 2;
				rsInstitutionAddress.LockType = 3;
				rsInstitutionAddress.Open();
				while (!rsInstitutionAddress.EOF) {
					var rsProvince = Server.CreateObject("ADODB.Recordset");
					rsProvince.ActiveConnection = MM_cnnASP02_STRING;
					rsProvince.Source = "{call dbo.cp_prov_state2("+ rsInstitutionAddress("intprvst_id") + ",'','',0,1,'Q',0)}";
					rsProvince.CursorType = 0;
					rsProvince.CursorLocation = 2;
					rsProvince.LockType = 3;
					rsProvince.Open();

					To_Address = Trim(rsInstitutionAddress.Fields.Item("chvSchool_Name").Value) + "<br>" + Trim(String(rsInstitutionAddress.Fields.Item("chvAddress").Value).replace(/\n/,"<br>")) + "<br>" + Trim(rsInstitutionAddress("chvCity")) + ", " + Trim(rsProvince("chrprvst_abbv")) + "<br>" + Trim(rsInstitutionAddress("chvcntry_name")) + " " + FormatPostalCode(rsInstitutionAddress("chvPostal_zip"));
					rsInstitutionAddress.MoveNext();
				}
				rsInstitutionAddress.Close();
			break;
			case "13":
				var rsCompanyAddress = Server.CreateObject("ADODB.Recordset");
				rsCompanyAddress.ActiveConnection = MM_cnnASP02_STRING;
				rsCompanyAddress.Source = "{call dbo.cp_company_address("+ rsContact.Fields.Item("insWork_id").Value + ",0,'','',0,'',0,'','','',0,'','','',0,'','','','','',0,1,'Q',0)}";
				rsCompanyAddress.CursorType = 0;
				rsCompanyAddress.CursorLocation = 2;
				rsCompanyAddress.LockType = 3;
				rsCompanyAddress.Open();

				if (!rsCompanyAddress.EOF) {
					if (String(rsCompanyAddress.Fields.Item("intAddress_id").Value)>0) {
						To_Address = Trim(rsCompanyAddress("chvAddress")) + "<br>" + Trim(rsCompanyAddress("chvCity")) + ", " + Trim(rsCompanyAddress("chrprvst_abbv")) + "<br>" + Trim(rsCompanyAddress("chvcntry_name")) + " " + FormatPostalCode(rsCompanyAddress("chvPostal_zip"));
					}
				}
			break;
			case "14":
				var rsCompanyAddress = Server.CreateObject("ADODB.Recordset");
				rsCompanyAddress.ActiveConnection = MM_cnnASP02_STRING;
				rsCompanyAddress.Source = "{call dbo.cp_company_address("+ rsContact.Fields.Item("insWork_id").Value + ",0,'','',0,'',0,'','','',0,'','','',0,'','','','','',0,1,'Q',0)}";
				rsCompanyAddress.CursorType = 0;
				rsCompanyAddress.CursorLocation = 2;
				rsCompanyAddress.LockType = 3;
				rsCompanyAddress.Open();

				if (!rsCompanyAddress.EOF) {
					if (String(rsCompanyAddress.Fields.Item("intAddress_id").Value)>0) {
						To_Address = Trim(rsCompanyAddress("chvAddress")) + "<br>" + Trim(rsCompanyAddress("chvCity")) + ", " + Trim(rsCompanyAddress("chrprvst_abbv")) + "<br>" + Trim(rsCompanyAddress("chvcntry_name")) + " " + FormatPostalCode(rsCompanyAddress("chvPostal_zip"));
					}
				}
			break;
			default:
				var rsContactAddress = Server.CreateObject("ADODB.Recordset");
				rsContactAddress.ActiveConnection = MM_cnnASP02_STRING;
				rsContactAddress.Source = "{call dbo.cp_Contact_Address("+ Request.QueryString("To") + ",0,'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,0,'Q',0)}";
				rsContactAddress.CursorType = 0;
				rsContactAddress.CursorLocation = 2;
				rsContactAddress.LockType = 3;
				rsContactAddress.Open();
				while (!rsContactAddress.EOF) {
					if (Trim(rsContactAddress("chvAddrs_type"))=="W") {
						To_Address = Trim(rsContactAddress("chvAddress")) + "<br>" + Trim(rsContactAddress("chvCity")) + ", " + Trim(rsContactAddress("chvProv")) + "<br>" + Trim(rsContactAddress("chvCountry")) + " " + FormatPostalCode(rsContactAddress("chvPostal_zip"));
					}
					rsContactAddress.MoveNext();
				}
				rsContactAddress.Close();
			break;
		}						
	}
	rsContact.Close();
} else {
	var rsClient = Server.CreateObject("ADODB.Recordset");
	rsClient.ActiveConnection = MM_cnnASP02_STRING;
	rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("To") + ")}";
	rsClient.CursorType = 0;
	rsClient.CursorLocation = 2;
	rsClient.LockType = 3;
	rsClient.Open();

	if (!rsClient.EOF) {
		var To_Whole_Name = rsClient("chvFst_Name") + " " + rsClient("chvLst_Name") + "<br>";	
		var rsClientAddress = Server.CreateObject("ADODB.Recordset");
		rsClientAddress.ActiveConnection = MM_cnnASP02_STRING;
		rsClientAddress.Source = "{call dbo.cp_Adult_Address("+ Request.QueryString("To") + ")}";
		rsClientAddress.CursorType = 0;
		rsClientAddress.CursorLocation = 2;
		rsClientAddress.LockType = 3;
		rsClientAddress.Open();
		if (!rsClientAddress.EOF) {
				To_Address = Trim(rsClientAddress("chvAddress")) + "<br>" + Trim(rsClientAddress("chvCity")) + ", " + Trim(rsClientAddress("chvProv")) + "<br>" + Trim(rsClientAddress("chvCountry")) + " " + FormatPostalCode(rsClientAddress("chvPostal_zip"));		
				rsClientAddress.MoveNext();
		}
		
		while (!rsClientAddress.EOF) {
			if (Trim(rsClientAddress("chvAddrs_type"))=="A") {
				To_Address = Trim(rsClientAddress("chvAddress")) + "<br>" + Trim(rsClientAddress("chvCity")) + ", " + Trim(rsClientAddress("chvProv")) + "<br>" + Trim(rsClientAddress("chvCountry")) + " " + FormatPostalCode(rsClientAddress("chvPostal_zip"));
			}				
			rsClientAddress.MoveNext();
		}
		rsClientAddress.Close();
	}
	rsClient.Close();
}
%>
<html>
<body bgcolor="#FFFFFF" text="#000000">
<table cellpadding="2" cellspacing="3" align="top">
	<tr>
		<td valign="top" width="450" style="font-size: 12pt; font-weight: bold">
			Sirius Innovations Inc.<br>
			P.O. Box 43119 Richmond Ctr PO <br>
			Richmond, B.C. V6V 2W4 </td>
		<td></td>
	</tr>
	<tr>
		<td></td>
		<td style="font-size: 12pt; font-weight: bold">
			<br><br><br>
			<%=To_Whole_Name%>
			<%=To_Address%>			
		</td>
	</tr>
</table>
</body>
</html>
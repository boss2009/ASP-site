<%
var Creation_Date = "";

var Sender_Email = "";

var Recipient_Title = "";
var Recipient_Whole_Name = "";
var Recipient_First_Name = "";
var Recipient_Last_Name = "";
var Recipient_Work_Address = "";
var Recipient_Job_Position = "";
var Recipient_Email = "";

var Subject_Whole_Name = "";
var Subject_First_Name = "";
var Subject_Last_Name = "";
var Subject_SIN = "";

var Loaned_Equipment_List = "";

var Loan_Conditions = "";

var Document_Conditions = "";

var Trainer = "";
var Trainer_Work_Phone = "";
var Training_Requested = ""

var CC_Whole_Name = "";
var CC_Email = "";

if (Request.Form("Mode") == "0") {
	Creation_Date = "{Creation_Date}";

	Sender_Email = "{Sender_Email}";

	Recipient_Title = "{Recipient_Title}";
	Recipient_Whole_Name = "{Recipient_Whole_Name}";
	Recipient_First_Name = "{Recipient_First_Name}";
	Recipient_Last_Name = "{Recipient_Last_Name}";
	Recipient_Work_Address = "{Recipient_Work_Address}";
	Recipient_Job_Position = "{Recipient_Job_Position}";
	Recipient_Email = "{Recipient_Email}";

	Subject_Whole_Name = "{Subject_Whole_Name}";
	Subject_First_Name = "{Subject_First_Name}";
	Subject_Last_Name = "{Subject_Last_Name}";
	Subject_SIN = "{Subject_SIN}";

	Loaned_Equipment_List = "{Loaned_Equipment_List}";

	Loan_Conditions = "{Loan_Conditions}";

	Document_Conditions = "{Document_Conditions}";

	Trainer = "{Trainer}";
	Trainer_Work_Phone = "{Trainer_Work_Phone}";
	Training_Requested = "{Training_Requested}";

	CC_Whole_Name = "{CC_Whole_Name}";
	CC_Email = "{CC_Email}";
} else {
	Creation_Date = Request.Form("DateGenerated")

	var rsContact = Server.CreateObject("ADODB.Recordset");
	rsContact.ActiveConnection = MM_cnnASP02_STRING;
	rsContact.Source = "{call dbo.cp_contacts("+Request.Form("Recipient")+",0,'','','',0,0,0,1,0,'',1,'Q',0)}"
	rsContact.CursorType = 0;
	rsContact.CursorLocation = 2;
	rsContact.LockType = 3;
	rsContact.Open();

	if (!rsContact.EOF) {
		Recipient_Title = ((rsContact("insTitle_Typ_id") > 0)?Trim(rsContact("chvtitle")):"");
		Recipient_First_Name = Trim(rsContact("chvFst_Name"));
		Recipient_Last_Name = Trim(rsContact("chvLst_Name"));
		Recipient_Whole_Name = Recipient_First_Name + " " + Recipient_Last_Name;

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

					Recipient_Work_Address = Trim(rsInstitutionAddress.Fields.Item("chvSchool_Name").Value) + "<br>" + Trim(String(rsInstitutionAddress.Fields.Item("chvAddress").Value).replace(/\n/,"<br>")) + "<br>" + Trim(rsInstitutionAddress("chvCity")) + " " + Trim(rsProvince("chrprvst_abbv")) + " " + FormatPostalCode(rsInstitutionAddress("chvPostal_zip"));
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
						Recipient_Work_Address = Trim(rsCompanyAddress("chvAddress")) + "<br>" + Trim(rsCompanyAddress("chvCity")) + " " + Trim(rsCompanyAddress("chrprvst_abbv")) + " " + FormatPostalCode(rsCompanyAddress("chvPostal_zip"));
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
						Recipient_Work_Address = Trim(rsCompanyAddress("chvAddress")) + "<br>" + Trim(rsCompanyAddress("chvCity")) + " " + Trim(rsCompanyAddress("chrprvst_abbv")) + " " + FormatPostalCode(rsCompanyAddress("chvPostal_zip"));
					}
				}
			break;
			default:
				var rsContactAddress = Server.CreateObject("ADODB.Recordset");
				rsContactAddress.ActiveConnection = MM_cnnASP02_STRING;
				rsContactAddress.Source = "{call dbo.cp_Contact_Address("+ Request.Form("Recipient") + ",0,'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,0,'Q',0)}";
				rsContactAddress.CursorType = 0;
				rsContactAddress.CursorLocation = 2;
				rsContactAddress.LockType = 3;
				rsContactAddress.Open();
				while (!rsContactAddress.EOF) {
					if (Trim(rsContactAddress("chvAddrs_type"))=="W") {
						Recipient_Work_Address = Trim(rsContactAddress("chvAddress")) + "<br>" + Trim(rsContactAddress("chvCity")) + " " + Trim(rsContactAddress("chvProv")) + FormatPostalCode(rsContactAddress("chvPostal_zip"));
					}
					rsContactAddress.MoveNext();
				}
				rsContactAddress.Close();
			break;
		}

		var rsContactAddress = Server.CreateObject("ADODB.Recordset");
		rsContactAddress.ActiveConnection = MM_cnnASP02_STRING;
		rsContactAddress.Source = "{call dbo.cp_Contact_Address("+ Request.Form("Recipient") + ",0,'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,0,'Q',0)}";
		rsContactAddress.CursorType = 0;
		rsContactAddress.CursorLocation = 2;
		rsContactAddress.LockType = 3;
		rsContactAddress.Open();
		while (!rsContactAddress.EOF) {
			if (Trim(rsContactAddress("chvAddrs_type"))=="W") {
				Recipient_Email = Trim(rsContactAddress("chvemail"));
			}
			rsContactAddress.MoveNext();
		}
		rsContactAddress.Close();

		Recipient_Job_Position = Trim(rsContact("chvJob_Title"));
	}

	rsContact.Close();

	var rsClient = Server.CreateObject("ADODB.Recordset");
	rsClient.ActiveConnection = MM_cnnASP02_STRING;
	rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.Form("Subject") + ")}";
	rsClient.CursorType = 0;
	rsClient.CursorLocation = 2;
	rsClient.LockType = 3;
	rsClient.Open();

	if (!rsClient.EOF) {
		Subject_First_Name = Trim(rsClient("chvFst_Name"));
		Subject_Last_Name = Trim(rsClient("chvLst_Name"));
		Subject_Whole_Name = Subject_First_Name + " " + Subject_Last_Name;
		Subject_SIN = FormatSIN(rsClient("chrSIN_no"));
	}

	rsClient.Close();

	var rsInventoryLoaned = Server.CreateObject("ADODB.Recordset");
	rsInventoryLoaned.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryLoaned.Source = "{call dbo.cp_eqp_loaned(0,"+Request.QueryString("intLoan_Req_id")+",0,'',0,0,'','',0,'Q',0)}";
	rsInventoryLoaned.CursorType = 0;
	rsInventoryLoaned.CursorLocation = 2;
	rsInventoryLoaned.LockType = 3;
	rsInventoryLoaned.Open();

	if (!rsInventoryLoaned.EOF) {
		Loaned_Equipment_List = "<ul>";
	}
	while (!rsInventoryLoaned.EOF) {
		Loaned_Equipment_List = Loaned_Equipment_List + "<li>" + Trim(rsInventoryLoaned("chvInventory_Name"));
		rsInventoryLoaned.MoveNext();
	}
	if (Loaned_Equipment_List.length > 0) Loaned_Equipment_List = Loaned_Equipment_List + "</ul>";
	rsInventoryLoaned.Close();

	if ((Request.Form("LoanConditionOne") > "0") || (Request.Form("LoanConditionTwo") > "0") || (Request.Form("LoanConditionThree") > "0") || (Request.Form("LoanConditionFour") > "0") || (String(Request.Form("LoanConditionOther")) != "")) {
		var rsDocumentCondition = Server.CreateObject("ADODB.Recordset");
		rsDocumentCondition.ActiveConnection = MM_cnnASP02_STRING;
		rsDocumentCondition.CursorType = 0;
		rsDocumentCondition.CursorLocation = 2;
		rsDocumentCondition.LockType = 3;

		Loan_Conditions = "Please note the following condition(s) apply to this equipment loan:<br>";
		Loan_Conditions = Loan_Conditions + "<ul>";
		if (Request.Form("LoanConditionOne")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("LoanConditionOne") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			var tempdesc = Trim(rsDocumentCondition("chvDocDesc"));
			Loan_Conditions = Loan_Conditions + "<li>" + tempdesc.substring(0, tempdesc.length-4);
			rsDocumentCondition.Close();
		}
		if (Request.Form("LoanConditionTwo")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("LoanConditionTwo") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			var tempdesc = Trim(rsDocumentCondition("chvDocDesc"));
			Loan_Conditions = Loan_Conditions + "<li>" + tempdesc.substring(0, tempdesc.length-4);

			rsDocumentCondition.Close();
		}
		if (Request.Form("LoanConditionThree")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("LoanConditionThree") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			var tempdesc = Trim(rsDocumentCondition("chvDocDesc"));
			Loan_Conditions = Loan_Conditions + "<li>" + tempdesc.substring(0, tempdesc.length-4);

			rsDocumentCondition.Close();
		}
		if (Request.Form("LoanConditionFour")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("LoanConditionFour") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			var tempdesc = Trim(rsDocumentCondition("chvDocDesc"));
			Loan_Conditions = Loan_Conditions + "<li>" + tempdesc.substring(0, tempdesc.length-4);
			rsDocumentCondition.Close();
		}
		if (String(Trim(Request.Form("LoanConditionOther"))) != "") {
			Loan_Conditions = Loan_Conditions + "<li>" + Request.Form("LoanConditionOther");
		}
		Loan_Conditions = Loan_Conditions + "</ul>";
	}

	if ((Request.Form("DocumentConditionOne") > "0") || (Request.Form("DocumentConditionTwo") > "0") || (Request.Form("DocumentConditionThree") > "0") || (Request.Form("DocumentConditionFour") > "0") || (String(Request.Form("DocumentConditionOther")) != "")) {
		var rsDocumentCondition = Server.CreateObject("ADODB.Recordset");
		rsDocumentCondition.ActiveConnection = MM_cnnASP02_STRING;
		rsDocumentCondition.CursorType = 0;
		rsDocumentCondition.CursorLocation = 2;
		rsDocumentCondition.LockType = 3;

		Document_Conditions = "AT-BC personnel will contact " + Subject_First_Name + " to make arrangements for delivery of equipment when we receive the following:";
		Document_Conditions = Document_Conditions + "<ul>";
		if (Request.Form("DocumentConditionOne")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DocumentConditionOne") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			var tempdesc = Trim(rsDocumentCondition("chvDocDesc"));
			Document_Conditions = Document_Conditions + "<li>" + tempdesc.substring(0, tempdesc.length-4);
			rsDocumentCondition.Close();
		}
		if (Request.Form("DocumentConditionTwo")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DocumentConditionTwo") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			var tempdesc = Trim(rsDocumentCondition("chvDocDesc"));
			Document_Conditions = Document_Conditions + "<li>" + tempdesc.substring(0, tempdesc.length-4);
			rsDocumentCondition.Close();
		}
		if (Request.Form("DocumentConditionThree")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DocumentConditionThree") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			var tempdesc = Trim(rsDocumentCondition("chvDocDesc"));
			Document_Conditions = Document_Conditions + "<li>" + tempdesc.substring(0, tempdesc.length-4);
			rsDocumentCondition.Close();
		}
		if (Request.Form("DocumentConditionFour")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DocumentConditionFour") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			var tempdesc = Trim(rsDocumentCondition("chvDocDesc"));
			Document_Conditions = Document_Conditions + "<li>" + tempdesc.substring(0, tempdesc.length-4);
			rsDocumentCondition.Close();
		}
		if (String(Trim(Request.Form("DocumentConditionOther"))) != "") {
			Document_Conditions = Document_Conditions + "<li>" + Request.Form("DocumentConditionOther");
		}
		Document_Conditions = Document_Conditions + "</ul>";
	}

	if (Request.Form("TrainingRequested")=="on"){
		Training_Requested = "If you have adaptive equipment and require training, an AT-BC trainer will contact you to schedule the training.";
	}

	if ((Request.Form("CC")!="") || (Request.Form("CCClient")=="on")) {
		CC_Whole_Name = "<p>cc:<br>";
		if (Request.Form("CCClient")=="on") {
			var rsClient = Server.CreateObject("ADODB.Recordset");
			rsClient.ActiveConnection = MM_cnnASP02_STRING;
			rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
			rsClient.CursorType = 0;
			rsClient.CursorLocation = 2;
			rsClient.LockType = 3;
			rsClient.Open();

			if (!rsClient.EOF) {
				CC_Whole_Name = CC_Whole_Name + Trim(rsClient("chvFst_Name")) + " " + Trim(rsClient("chvLst_Name")) + "<br>";
			}

			rsClient.Close();
		}
		if (Request.Form("CC")!="") {
			tempCC = String(Request.Form("CC")).split(":");
			for (var i=0; i < tempCC.length; i++) {
				var rsContact = Server.CreateObject("ADODB.Recordset");
				rsContact.ActiveConnection = MM_cnnASP02_STRING;
				rsContact.Source = "{call dbo.cp_contacts("+tempCC[i]+",0,'','','',0,0,0,1,0,'',1,'Q',0)}"
				rsContact.CursorType = 0;
				rsContact.CursorLocation = 2;
				rsContact.LockType = 3;
				rsContact.Open();
				if (!rsContact.EOF) {
					CC_Whole_Name = CC_Whole_Name + rsContact("chvFst_Name") + " " + rsContact("chvLst_Name")
					if (Trim(rsContact("chvJob_Title"))!="") {
						CC_Whole_Name = CC_Whole_Name + ", " + rsContact("chvJob_Title");
					}
					var rsContactAddress = Server.CreateObject("ADODB.Recordset");
					rsContactAddress.ActiveConnection = MM_cnnASP02_STRING;
					rsContactAddress.Source = "{call dbo.cp_Contact_Address("+ tempCC[i] + ",0,'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,0,'Q',0)}";
					rsContactAddress.CursorType = 0;
					rsContactAddress.CursorLocation = 2;
					rsContactAddress.LockType = 3;
					rsContactAddress.Open();
					while (!rsContactAddress.EOF) {
						if (rsContactAddress("chvAddrs_type")=="W") {
							if ((String(rsContactAddress("chvemail")) != "null") && (Trim(rsContactAddress("chvemail")) != "") && (CheckEmail(rsContactAddress("chvemail")))) {
								CC_Email = CC_Email + Trim(rsContactAddress("chvemail")) + ";"
							}
						}
						rsContactAddress.MoveNext();
					}
					rsContactAddress.Close();
					if (rsContact("insWork_id") > "0") {
						var rsCompany = Server.CreateObject("ADODB.Recordset");
						rsCompany.ActiveConnection = MM_cnnASP02_STRING;
						rsCompany.Source = "{call dbo.cp_Company2("+rsContact("insWork_id")+",'',0,0,0,0,0,1,0,'',1,'Q',0)}"
						rsCompany.CursorType = 0;
						rsCompany.CursorLocation = 2;
						rsCompany.LockType = 3;
						rsCompany.Open();
						if (!rsCompany.EOF) {
							if (Trim(rsCompany("chvCompany_Name"))!="") {
								CC_Whole_Name = CC_Whole_Name + ", " + rsCompany("chvCompany_Name");
							}
						}
						rsCompany.Close();
					}
					CC_Whole_Name = CC_Whole_Name + "<br>";
				}
				rsContact.Close();
			}
		}
		CC_Whole_Name = CC_Whole_Name + "</p>";
	}


	var rsStaff = Server.CreateObject("ADODB.Recordset");
	rsStaff.ActiveConnection = MM_cnnASP02_STRING;
	rsStaff.Source = "{call dbo.cp_staff2("+Session("insStaff_id")+",0,'','',0,'','',0,0,0,0,0,0,0,0,0,1,0,'',1,'Q',0)}";
	rsStaff.CursorType = 0;
	rsStaff.CursorLocation = 2;
	rsStaff.LockType = 3;
	rsStaff.Open();

	if (!rsStaff.EOF) {
		var rsStaffAddress = Server.CreateObject("ADODB.Recordset");
		rsStaffAddress.ActiveConnection = MM_cnnASP02_STRING;
		rsStaffAddress.Source = "{call dbo.cp_staff_address(0,"+rsStaff("intAddress_id")+",'','',0,'',0,'','','',0,'','','',0,'','','','','',0,1,'Q',0)}";
		rsStaffAddress.CursorType = 0;
		rsStaffAddress.CursorLocation = 2;
		rsStaffAddress.LockType = 3;
		rsStaffAddress.Open();
		if (!rsStaffAddress.EOF) {
			Sender_Email = Trim(rsStaffAddress("chvEmail"));
			rsStaffAddress.Close();
		}
		rsStaff.Close();
	}
}
%>
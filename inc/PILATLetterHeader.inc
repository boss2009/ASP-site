<%
var Creation_Date = "";

var Sender_Email = "";

var Recipient_Title = "";
var Recipient_Whole_Name = "";
var Recipient_First_Name = "";
var Recipient_Last_Name = "";
var Recipient_Work_Address = "";
var Recipient_School_Address = "";
var Recipient_Job_Position = "";
var Recipient_Email = "";

var Pilat_Decline_Referral_Type = "";
var Decline_Reasons = "";

var Pilat_Accept_Referral_Type = "";
var Conditions = "";

var Introduction = "";

var Equipment_List = "";

var Equipment_Conditions = "";

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
	Recipient_School_Address = "{Recipient_School_Address}";
	Recipient_Job_Position = "{Recipient_Job_Position}";
	Recipient_Email = "{Recipient_Email}";

	Pilat_Decline_Referral_Type = "{Pilat_Decline_Referral_Type}";
	Decline_Reasons = "{Decline_Reasons}";

	Pilat_Accept_Referral_Type = "{Pilat_Accept_Referral_Type}";
	Conditions = "{Conditions}";

	Introduction = "{Introduction}";

	Equipment_List = "{Equipment_List}";
	Equipment_Conditions = "{Equipment_Conditions}";
	Document_Conditions = "{Document_Conditions}";

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

					Recipient_Work_Address = Trim(rsInstitutionAddress.Fields.Item("chvSchool_Name").Value) + "<br>" + Trim(String(rsInstitutionAddress.Fields.Item("chvAddress").Value).replace(/\n/,"<br>")) + "<br>" + Trim(rsInstitutionAddress("chvCity")) + " " + Trim(rsProvince("chrprvst_abbv")) + FormatPostalCode(rsInstitutionAddress("chvPostal_zip"));
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
						Recipient_Work_Address = Trim(rsContactAddress("chvAddress")) + "<br>" + Trim(rsContactAddress("chvCity")) + " " + Trim(rsContactAddress("chvProv")) + " " + FormatPostalCode(rsContactAddress("chvPostal_zip"));
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

	switch (String(Request.Form("Conditions"))) {
		case "1":
			Conditions = "<p>We are pleased to inform you that AT-BC is able to provide training on the specified adaptive technology owned by your institution.  An AT-BC trainer will contact you to schedule the training for your designated institutional staff that will be supporting the equipment.  "+Request.Form("OtherConditions")+"</p>";
			Conditions = Conditions + "<p>By providing this training, we hope to increase your staff's ability to support the institution's adaptive technology and the students that require access to this equipment to pursue their educational goals.";
		break;
		case "2":
			Conditions = "<p>We are pleased to inform you that AT-BC is able to provide consultation on your institution's adaptive technology needs.  "+Request.Form("OtherConditions")+"</p>";
			Conditions = Conditions + "<p>By consulting with your institution, we hope to provide the information your institution requires to enhance access to adaptive equipment that will enable your students with disabilities to pursue their educational goals more effectively.</p>";
		break;
	}

	switch (String(Request.Form("PilatDeclineReferralType"))) {
		case "1":
			Pilat_Decline_Referral_Type = "PILAT Referral for Low Utilization Loan";
		break;
		case "2":
			Pilat_Decline_Referral_Type = "PILAT Referral for Interim Loan";
		break;
		case "3":
			Pilat_Decline_Referral_Type = "PILAT Referral for Donation";
		break;
		case "4":
			Pilat_Decline_Referral_Type = "PILAT Referral for Equipment Purchase";
		break;
	}

	switch (String(Request.Form("PilatAcceptReferralType"))) {
		case "1":
			Introduction = "<p>We are please to inform you that your institution has been accepted for a low utilization loan of the following equipment:</p>";
			Pilat_Accept_Referral_Type = "PILAT Referral for Low Utilization Loan";
		break;
		case "2":
			Introduction = "<p>We are please to inform you that your institution has been accepted for an interim loan of the following equipment:</p>";
			Pilat_Accept_Referral_Type = "PILAT Referral for Interim Loan";
		break;
		case "3":
			Introduction = "<p>We are please to inform you that the Assistive Technology - British Columbia will be donating the following equipment to your institution:</p>";
			Pilat_Accept_Referral_Type = "PILAT Referral for Donation";
		break;
		case "4":
			Introduction = "<p>We are please to inform you that the Assistive Technology - British Columbia has agreed to purchase the following equipment on your institution's behalf:</p>";
			Pilat_Accept_Referral_Type = "PILAT Referral for Equipment Purchase";
		break;
		case "5":
			Pilat_Accept_Referral_Type = "Training";
		break;
		case "6":
			Pilat_Accept_Referral_Type = "Consultation";
		break;
	}

	if (Request.Form("TransactionType")=="Buyout") {
		var rsInventorySold = Server.CreateObject("ADODB.Recordset");
		rsInventorySold.ActiveConnection = MM_cnnASP02_STRING;
		rsInventorySold.Source = "{call dbo.cp_buyout_eqp_sold(0,"+Request.QueryString("intBuyout_req_id")+",0,0.0,'',0,0,'',0,'Q',0)}";
		rsInventorySold.CursorType = 0;
		rsInventorySold.CursorLocation = 2;
		rsInventorySold.LockType = 3;
		rsInventorySold.Open();

		if (!rsInventorySold.EOF) {
			Equipment_List = "<ul>";
		}
		while (!rsInventorySold.EOF) {
			Equipment_List = Equipment_List + "<li>" + Trim(rsInventorySold("chvInventory_Name"));
			rsInventorySold.MoveNext();
		}
		if (Trim(Request.Form("OtherEquipmentList"))!="") {
			Equipment_List = Equipment_List + "<li>" + Trim(Request.Form("OtherEquipmentList"));
		}
		if (Equipment_List.length > 0)	Equipment_List = Equipment_List + "</ul>";
		rsInventorySold.Close();
	}
	if (Request.Form("TransactionType")=="Loan") {
		var rsInventoryLoaned = Server.CreateObject("ADODB.Recordset");
		rsInventoryLoaned.ActiveConnection = MM_cnnASP02_STRING;
		rsInventoryLoaned.Source = "{call dbo.cp_eqp_loaned(0,"+Request.QueryString("intLoan_Req_id")+",0,'',0,0,'','',0,'Q',0)}";
		rsInventoryLoaned.CursorType = 0;
		rsInventoryLoaned.CursorLocation = 2;
		rsInventoryLoaned.LockType = 3;
		rsInventoryLoaned.Open();

		if (!rsInventoryLoaned.EOF) {
			Equipment_List = "<ul>";
		}
		while (!rsInventoryLoaned.EOF) {
			Equipment_List = Equipment_List + "<li>" + Trim(rsInventoryLoaned("chvInventory_Name"));
			rsInventoryLoaned.MoveNext();
		}
		if (Trim(Request.Form("OtherEquipmentList"))!="") {
			Equipment_List = Equipment_List + "<li>" + Trim(Request.Form("OtherEquipmentList"));
		}
		if (Equipment_List.length > 0) Equipment_List = Equipment_List + "</ul>";
		rsInventoryLoaned.Close();
	}

	switch (String(Request.Form("EquipmentConditions"))) {
		case "1":
			Equipment_Conditions = "<p>The purpose of this Low Utilization Loan is to enhance access at the institution in";
			Equipment_Conditions += " situations where purchase of the adaptive equipment would be unreasonable due";
  			Equipment_Conditions += " to the low need for this specialized equipment.  As such, the loan equipment must";
  			Equipment_Conditions += " be returned to AT-BC when the student is no longer enrolled at the institution or no";
  			Equipment_Conditions += " longer requires access to this equipment at the institution.  AT-BC will review the";
  			Equipment_Conditions += " continued need for the equipment loan on " + Request.Form("LoanReviewDate") + " to ensure PILAT criteria are being met.</p>"
		break;
		case "2":
			Equipment_Conditions = "<p>The purpose of this Interim Loan is to ensure the availability of adaptive";
			Equipment_Conditions += " equipment for students while the institution is in the process of acquiring the";
			Equipment_Conditions += " specialized equipment.  As such, the loan equipment is expected to be returned on " + Request.Form("ReturnDate") + ".</p>";
		break;
		case "3":
			Equipment_Conditions = "<p>This equipment is being donated in as-is condition and AT-BC cannot be responsible for future technical support.</p>"
		break;
		case "4":
			Equipment_Conditions = "<p>The equipment purchase plan is as follows:";
			Equipment_Conditions += "<ul><li>AT-BC to invoice the institution for the equipment ordered on its behalf.";
			Equipment_Conditions += "<li>AT-BC to install and configure the software and hardware as applicable.";
			Equipment_Conditions += "<li>Institution staff will be responsible for on-going technical support and warranty service.";
			if (Trim(Request.Form("OtherEquipmentConditions"))!="") {
				Equipment_Conditions += "<li>" + Request.Form("OtherEquipmentConditions");
			}
			Equipment_Conditions += "</ul></p>";
		break;
	}


	if ((Request.Form("DeclineReasonOne") > "0") || (Request.Form("DeclineReasonTwo") > "0") || (Request.Form("DeclineReasonThree") > "0") || (Request.Form("DeclineReasonFour") > "0") || (String(Request.Form("OtherDeclineReason")) != "")) {
		var rsDeclineReason = Server.CreateObject("ADODB.Recordset");
		rsDeclineReason.ActiveConnection = MM_cnnASP02_STRING;
		rsDeclineReason.CursorType = 0;
		rsDeclineReason.CursorLocation = 2;
		rsDeclineReason.LockType = 3;

		Decline_Reasons = "<ul>";
		if (Request.Form("DeclineReasonOne")>"0") {
			rsDeclineReason.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DeclineReasonOne") + ",'','',1,'Q',0)}";
			rsDeclineReason.Open();
			Decline_Reasons = Decline_Reasons + "<li>" + rsDeclineReason("chvDocDesc");
			rsDeclineReason.Close();
		}
		if (Request.Form("DeclineReasonTwo")>"0") {
			rsDeclineReason.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DeclineReasonTwo") + ",'','',1,'Q',0)}";
			rsDeclineReason.Open();
			Decline_Reasons = Decline_Reasons + "<li>" + rsDeclineReason("chvDocDesc");
			rsDeclineReason.Close();
		}
		if (Request.Form("DeclineReasonThree")>"0") {
			rsDeclineReason.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DeclineReasonThree") + ",'','',1,'Q',0)}";
			rsDeclineReason.Open();
			Decline_Reasons = Decline_Reasons + "<li>" + rsDeclineReason("chvDocDesc");
			rsDeclineReason.Close();
		}
		if (Request.Form("DeclineReasonFour")>"0") {
			rsDeclineReason.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DeclineReasonFour") + ",'','',1,'Q',0)}";
			rsDeclineReason.Open();
			Decline_Reasons = Decline_Reasons + "<li>" + rsDeclineReason("chvDocDesc");
			rsDeclineReason.Close();
		}
		if (String(Trim(Request.Form("OtherDeclineReason"))) != "") {
			Decline_Reasons = Decline_Reasons + "<li>" + Request.Form("OtherDeclineReason");
		}
		Decline_Reasons = Decline_Reasons + "</ul>";
	}

	if ((Request.Form("DocumentConditionOne") > "0") || (Request.Form("DocumentConditionTwo") > "0") || (Request.Form("DocumentConditionThree") > "0") || (Request.Form("DocumentConditionFour") > "0") || (String(Request.Form("OtherDocumentCondition")) != "")) {
		var rsDocumentCondition = Server.CreateObject("ADODB.Recordset");
		rsDocumentCondition.ActiveConnection = MM_cnnASP02_STRING;
		rsDocumentCondition.CursorType = 0;
		rsDocumentCondition.CursorLocation = 2;
		rsDocumentCondition.LockType = 3;

		Document_Conditions = "The equipment request is now being processed and will be shipped as soon as possible when we receive the following outstanding information:";
		Document_Conditions = Document_Conditions + "<ul>";
		if (Request.Form("DocumentConditionOne")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DocumentConditionOne") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			Document_Conditions = Document_Conditions + "<li>" + rsDocumentCondition("chvDocDesc");
			rsDocumentCondition.Close();
		}
		if (Request.Form("DocumentConditionTwo")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DocumentConditionTwo") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			Document_Conditions = Document_Conditions + "<li>" + rsDocumentCondition("chvDocDesc");
			rsDocumentCondition.Close();
		}
		if (Request.Form("DocumentConditionThree")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DocumentConditionThree") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			Document_Conditions = Document_Conditions + "<li>" + rsDocumentCondition("chvDocDesc");
			rsDocumentCondition.Close();
		}
		if (Request.Form("DocumentConditionFour")>"0") {
			rsDocumentCondition.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("DocumentConditionFour") + ",'','',1,'Q',0)}";
			rsDocumentCondition.Open();
			Document_Conditions = Document_Conditions + "<li>" + rsDocumentCondition("chvDocDesc");
			rsDocumentCondition.Close();
		}
		if (String(Trim(Request.Form("OtherDocumentCondition"))) != "") {
			Document_Conditions = Document_Conditions + "<li>" + Request.Form("OtherDocumentCondition");
		}
		Document_Conditions = Document_Conditions + "</ul>";
	} else {
		Document_Conditions = "<p>The equipment request is now being processed and will be shipped as soon as possible.</p>";
	}

	if (Request.Form("TrainingRequested")=="on"){
		Training_Requested = "To schedule training on the adaptive equipment, please contact AT-BC";
	}
	if (Request.Form("CC")!="") {
		CC_Whole_Name = "<p>cc:<br>";
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
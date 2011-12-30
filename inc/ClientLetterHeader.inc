<%
var Creation_Date = "";

var Sender_Email = "";

var Recipient_Title = "";
var Recipient_Whole_Name = "";
var Recipient_First_Name = "";
var Recipient_Last_Name = "";
var Recipient_Work_Address = "";
var Recipient_School_Address = "";
var Recipient_SIN = "";
var Recipient_Job_Position = "";
var Recipient_Email = "";

var Subject_Whole_Name = "";
var Subject_First_Name = "";
var Subject_Last_Name = "";
var Subject_SIN = "";

var Purpose_Of_Loan = "";
var Equipment_Still_On_Loan_List = "";
var Equipment_Still_On_Loan_List_With_Cost = "";
var Equipment_Still_On_Loan_List_Total_Cost = "";
var Equipment_Delivered_Date = "";
var Reason_For_Ineligibility = "";
var Purchase_Cost_Of_Equipment = "";

var Last_Loan_Default_Letter = "";
var Reason_For_Canceling_Default = "";

var Is_Employment_Buyout = "";
var Employment_Loan_Duration = "";
var Loan_Expiry_Date = "";
var Discount_Amount = "";
var Buyout_Cost = "";
var Number_Of_Installments = "";
var Installment_Due_Dates = "";
var Payment_In_Full_Date = "";

var CSG_Issues = "";
var CSG_Missing_Documentation = "";

var Loan_Issues = "";
var Loan_Missing_Documentation = "";

var Conditions_To_Maintain_Loan = "";
var Reply_By_Date = "";

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
	Recipient_SIN = "{Recipient_SIN}";
	Recipient_Email = "{Recipient_Email}";

	Subject_Whole_Name = "{Subject_Whole_Name}";
	Subject_First_Name = "{Subject_First_Name}";
	Subject_Last_Name = "{Subject_Last_Name}";
	Subject_SIN = "{Subject_SIN}";

	Purpose_Of_Loan = "{Purpose_Of_Loan}";
	Equipment_Still_On_Loan_List = "{Equipment_Still_On_Loan_List}";
	Equipment_Still_On_Loan_List_With_Cost = "{Equipment_Still_On_Loan_List_With_Cost}";
	Equipment_Still_On_Loan_List_Total_Cost = "{Equipment_Still_On_Loan_List_Total_Cost}";
	Equipment_Delivered_DAte = "{Equipment_Delivered_DAte}";
	Reason_For_Ineligibility = "{Reason_For_Ineligibility}";
	Purchase_Cost_Of_Equipment = "{Purchase_Cost_Of_Equipment}";

	Last_Loan_Default_Letter = "{Last_Loan_Default_Letter}";
	Reason_For_Canceling_Default = "{Reason_For_Canceling_Default}";

	Is_Employment_Buyout = "{Is_Employment_Buyout}";
	Employment_Loan_Duration = "{Employment_Loan_Duration}";
	Loan_Expiry_Date = "{Loan_Expiry_Date}";
	Discount_Amount = "{Discount_Amount}";
	Buyout_Cost = "{Buyout_Cost}";
	Number_Of_Installments = "{Number_Of_Installments}";
	Installment_Due_Dates = "{Installment_Due_Dates}";
	Payment_In_Full_Date = "{Payment_In_Full_Date}";

	CSG_Issues = "{CSG_Issues}";
	CSG_Missing_Documentation = "{CSG_Missing_Documentation}";

	Conditions_To_Maintain_Loan = "{Conditions_To_Maintain_Loan}";
	Reply_By_Date = "{Reply_By_Date}";

	CC_Whole_Name = "{CC Whole Name}";
	CC_Email = "{CC_Email}";
} else {
	Creation_Date = Request.Form("DateGenerated")

	var temp = String(Request.Form("Recipient")).split(":");

	if (temp[0] == "Contact") {
		var rsContact = Server.CreateObject("ADODB.Recordset");
		rsContact.ActiveConnection = MM_cnnASP02_STRING;
		rsContact.Source = "{call dbo.cp_contacts("+temp[1]+",0,'','','',0,0,0,1,0,'',1,'Q',0)}"
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
							Recipient_Work_Address = Trim(rsCompanyAddress("chvAddress")) + "<br>" + Trim(rsCompanyAddress("chvCity")) + " " + Trim(rsCompanyAddress("chrprvst_abbv")) + "<br>" + FormatPostalCode(rsCompanyAddress("chvPostal_zip"));
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
					rsContactAddress.Source = "{call dbo.cp_Contact_Address("+ temp[1] + ",0,'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,0,'Q',0)}";
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
			rsContactAddress.Source = "{call dbo.cp_Contact_Address("+ temp[1] + ",0,'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,0,'Q',0)}";
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
	} else {
		var rsClient = Server.CreateObject("ADODB.Recordset");
		rsClient.ActiveConnection = MM_cnnASP02_STRING;
		rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ temp[1] + ")}";
		rsClient.CursorType = 0;
		rsClient.CursorLocation = 2;
		rsClient.LockType = 3;
		rsClient.Open();

		if (!rsClient.EOF) {
			Recipient_Title = ((rsClient("bitGender_is_male") == "1")?"Mr.":"Ms.");
			Recipient_First_Name = Trim(rsClient("chvFst_Name"));
			Recipient_Last_Name = Trim(rsClient("chvLst_Name"));
			Recipient_Whole_Name = Recipient_First_Name + " " + Recipient_Last_Name;
			Recipient_SIN = FormatSIN(rsClient("chrSIN_no"))
			var rsClientAddress = Server.CreateObject("ADODB.Recordset");
			rsClientAddress.ActiveConnection = MM_cnnASP02_STRING;
			rsClientAddress.Source = "{call dbo.cp_Adult_Address("+ temp[1] + ")}";
			rsClientAddress.CursorType = 0;
			rsClientAddress.CursorLocation = 2;
			rsClientAddress.LockType = 3;
			rsClientAddress.Open();
			while (!rsClientAddress.EOF) {
				if (Trim(rsClientAddress("chvAddrs_type"))=="A") Recipient_School_Address = Trim(rsClientAddress("chvAddress")) + "<br>" + Trim(rsClientAddress("chvCity")) + " " + Trim(rsClientAddress("chvProv")) + " " + FormatPostalCode(rsClientAddress("chvPostal_zip"));
				Recipient_Email = Trim(rsClientAddress("chvemail"));
				rsClientAddress.MoveNext();
			}
			rsClientAddress.Close();
		}
		rsClient.Close();

	}

	if (Request.Form("Subject") != 0) {
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
	}

	var client = 0;
	if (Request.Form("Subject")==0) {
		client = temp[1];
	} else {
		client = Request.Form("Subject");
	}

	var rsEquipmentStillOnLoan = Server.CreateObject("ADODB.Recordset");
	rsEquipmentStillOnLoan.ActiveConnection = MM_cnnASP02_STRING;
	rsEquipmentStillOnLoan.Source = "{call dbo.cp_Loan_Request("+client+")}";
	rsEquipmentStillOnLoan.CursorType = 0;
	rsEquipmentStillOnLoan.CursorLocation = 2;
	rsEquipmentStillOnLoan.LockType = 3;
	rsEquipmentStillOnLoan.Open();

	//most recent loan with equip still on loan
	var first = 0;
	if (!rsEquipmentStillOnLoan.EOF) {
		Equipment_Still_On_Loan_List = "<ul>";
		Equipment_Still_On_Loan_List_With_Cost = "<ul>";
		Equipment_Still_On_Loan_List_Total_Cost = 0;
	}

	var temp = 0;
	while (!rsEquipmentStillOnLoan.EOF) {
		temp = 0
		if ((String(rsEquipmentStillOnLoan("dtsReturn"))=="null") && (String(rsEquipmentStillOnLoan("dtsDate_Shipped"))!="null")) {
			Equipment_Still_On_Loan_List = Equipment_Still_On_Loan_List + "<li>" + Trim(rsEquipmentStillOnLoan("chvInventory_Name"));
			Equipment_Still_On_Loan_List_With_Cost = Equipment_Still_On_Loan_List_With_Cost + "<li>" + Trim(rsEquipmentStillOnLoan("chvInventory_Name")) + "  (" + FormatCurrency(rsEquipmentStillOnLoan("fltPurchase_Cost")) + ")";
			temp = Number(rsEquipmentStillOnLoan("fltPurchase_Cost"));
			if (first == 0) first = rsEquipmentStillOnLoan.Fields.Item("intLoan_Req_id").Value;
		}
		rsEquipmentStillOnLoan.MoveNext();
		Equipment_Still_On_Loan_List_Total_Cost = Equipment_Still_On_Loan_List_Total_Cost + temp;
	}
	if (Equipment_Still_On_Loan_List.length > 0) Equipment_Still_On_Loan_List = Equipment_Still_On_Loan_List + "</ul>";
	if (Equipment_Still_On_Loan_List_With_Cost.length > 0) Equipment_Still_On_Loan_List_With_Cost = Equipment_Still_On_Loan_List_With_Cost + "</ul>";
	Equipment_Still_On_Loan_List_Total_Cost = FormatCurrency(Equipment_Still_On_Loan_List_Total_Cost);
	rsEquipmentStillOnLoan.Close();

	if (first != 0) {
		var rsLoan = Server.CreateObject("ADODB.Recordset");
		rsLoan.ActiveConnection = MM_cnnASP02_STRING;
		rsLoan.Source = "{call dbo.cp_loan_request2("+ first + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
		rsLoan.CursorType = 0;
		rsLoan.CursorLocation = 2;
		rsLoan.LockType = 3;
		rsLoan.Open();

		var intShip_dtl_id = 0;
		if (!rsLoan.EOF) {
			if (rsLoan.Fields.Item("intShip_dtl_id").Value != null) intShip_dtl_id = rsLoan.Fields.Item("intShip_dtl_id").Value;
			var rsMethod = Server.CreateObject("ADODB.Recordset");
			rsMethod.ActiveConnection = MM_cnnASP02_STRING;
			rsMethod.Source = "{call dbo.cp_loan_ship_method("+ intShip_dtl_id + ",0,'',0,0,0,'',0,'','',0,0,'Q',0)}";
			rsMethod.CursorType = 0;
			rsMethod.CursorLocation = 2;
			rsMethod.LockType = 3;
			rsMethod.Open();
			if (!rsMethod.EOF) {
				Equipment_Delivered_Date = FilterDate(rsMethod.Fields.Item("dtsDlvy_date").Value);
			}
			rsMethod.Close();
		}
		rsLoan.Close();
	}

	var rsPurposeOfLoan = Server.CreateObject("ADODB.Recordset");
	rsPurposeOfLoan.ActiveConnection = MM_cnnASP02_STRING;
	rsPurposeOfLoan.CursorType = 0;
	rsPurposeOfLoan.CursorLocation = 2;
	rsPurposeOfLoan.LockType = 3;
	rsPurposeOfLoan.Source = "{call dbo.cp_doc_cdn_rsn("+ Request.Form("PurposeOfLoan") + ",'','',1,'Q',0)}";
	rsPurposeOfLoan.Open();
	if (!rsPurposeOfLoan.EOF) {
		var tempdesc = String(Trim(rsPurposeOfLoan("chvDocDesc")));
		Purpose_Of_Loan = Purpose_Of_Loan + tempdesc.substring(0, tempdesc.length - 4);
	}
	rsPurposeOfLoan.Close();
	if (Trim(Request.Form("OtherPurposeOfLoan"))!="") Purpose_Of_Loan = Trim(Request.Form("OtherPurposeOfLoan"));

	var rsCorrespondence = Server.CreateObject("ADODB.Recordset");
	rsCorrespondence.ActiveConnection = MM_cnnASP02_STRING;
	rsCorrespondence.Source = "{call dbo.cp_get_ac_crsp_hstry2("+client+",0)}";
	rsCorrespondence.CursorType = 0;
	rsCorrespondence.CursorLocation = 2;
	rsCorrespondence.LockType = 3;
	rsCorrespondence.Open();
	while (!rsCorrespondence.EOF) {
		if (rsCorrespondence("insTemplate_id")==862) {
			Last_Loan_Default_Letter = FilterDate(rsCorrespondence("dtsSend_Date"));
			break;
		}
		rsCorrespondence.MoveNext();
	}

	switch (String(Request.Form("ReasonForCancelingDefault"))) {
		case "0":
			Reason_For_Canceling_Default = Trim(Request.Form("OtherReasonForCancelingDefault"));
		break;
		case "1":
			Reason_For_Canceling_Default = "the outstanding equipment that was on loan and have returned it to the AT-BC Loan Bank"
		break;
		case "2":
			Reason_For_Canceling_Default = "payment for the outstanding equipment that was on loan"
		break;
		default:
			Reason_For_Canceling_Default = Trim(Request.Form("OtherReasonForCancelingDefault"));
		break;
	}


	if ((Request.Form("ReasonForIneligibilityOne") > "0") || (Request.Form("ReasonForIneligibilityTwo") > "0") || (Request.Form("ReasonForIneligibilityThree") > "0") || (Request.Form("ReasonForIneligibilityFour") > "0") || (String(Request.Form("OtherReasonForIneligibility")) != "")) {
		Reason_For_Ineligibility = "<ul>";
		switch (String(Request.Form("ReasonForIneligibilityOne"))) {
			case "104":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " is not enrolled in an approved training or post secondary institution.";
			break;
			case "106":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " is not enrolled in required number of courses.";
			break;
			case "107":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " owes an outstanding payment for equipment purchase.";
			break;
			case "108":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>AT-BC has not been able to contact " + Subject_First_Name + " to discuss purchase plan or return of equipment.";
			break;
			case "109":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + "'s refusal to initiate purchase plan or return the loaned equipment.";
			break;
		}
		switch (String(Request.Form("ReasonForIneligibilityTwo"))) {
			case "104":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " is not enrolled in an approved training or post secondary institution.";
			break;
			case "106":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " is not enrolled in required number of courses.";
			break;
			case "107":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " owes an outstanding payment for equipment purchase.";
			break;
			case "108":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>AT-BC has not been able to contact " + Subject_First_Name + " to discuss purchase plan or return of equipment.";
			break;
			case "109":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + "'s refusal to initiate purchase plan or return the loaned equipment.";
			break;
		}
		switch (String(Request.Form("ReasonForIneligibilityThree"))) {
			case "104":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " is not enrolled in an approved training or post secondary institution.";
			break;
			case "106":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " is not enrolled in required number of courses.";
			break;
			case "107":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " owes an outstanding payment for equipment purchase.";
			break;
			case "108":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>AT-BC has not been able to contact " + Subject_First_Name + " to discuss purchase plan or return of equipment.";
			break;
			case "109":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + "'s refusal to initiate purchase plan or return the loaned equipment.";
			break;
		}
		switch (String(Request.Form("ReasonForIneligibilityFour"))) {
			case "104":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " is not enrolled in an approved training or post secondary institution.";
			break;
			case "106":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " is not enrolled in required number of courses.";
			break;
			case "107":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + " owes an outstanding payment for equipment purchase.";
			break;
			case "108":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>AT-BC has not been able to contact " + Subject_First_Name + " to discuss purchase plan or return of equipment.";
			break;
			case "109":
				Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Subject_First_Name + "'s refusal to initiate purchase plan or return the loaned equipment.";
			break;
		}
		if (Trim(Request.Form("OtherReasonForIneligibility")) != "") {
			Reason_For_Ineligibility = Reason_For_Ineligibility + "<li>" + Request.Form("OtherReasonForIneligibility") + ".";
		}
		Reason_For_Ineligibility = Reason_For_Ineligibility + "</ul>";
	}

	Purchase_Cost_Of_Equipment = FormatCurrency(Request.Form("PurchaseCostOfEquipment"));

	Is_Employment_Buyout = ((Request.Form("BuyoutType")=="0")?true:false);
	switch (String(Request.Form("EmploymentLoanDuration"))) {
		case "1":
			Employment_Loan_Duration = "1 year";
		break;
		case "2":
			Employment_Loan_Duration = "8 months";
		break;
		case "3":
			Employment_Loan_Duration = "6 months";
		break;
	}
	Loan_Expiry_Date = Request.Form("LoanExpiryDate");
	Discount_Amount = FormatCurrency(Request.Form("DiscountAmount"));
	Buyout_Cost = FormatCurrency(Request.Form("BuyoutCost"));
	Number_Of_Installments = Request.Form("NumberOfInstallments");
	if (Number_Of_Installments >= 1) {
		var temp = String(Request.Form("InstallmentDueDates")).split(":");
		for (var i = 0; i < temp.length; i++) {
			Installment_Due_Dates = Installment_Due_Dates + "Installment #" + (Number(i)+Number(1)) + " Due Date:" + temp[i] + "<br>";
		}
	}
	Payment_In_Full_Date = Request.Form("PaymentInFullDate");

	if ((Request.Form("NoBCSAPOrHNPT") == "on") || (Request.Form("BCSAPOrHNPTErrors") == "on") || (Request.Form("NoFinancialNeed") == "on") || (Request.Form("CanadaStudentLoanDefault") == "on") || (Request.Form("InsufficientTimeForProcessingTSSP") == "on") || (Request.Form("OutstandingReceipts") == "on") || (Request.Form("ATBCDefaultForBuyout") == "on") || (Request.Form("ContactATBCForClarifyEquipment") == "on") || (Request.Form("RequestForSecondSystem") == "on") || (Request.Form("IneligibleEquipment") == "on") || (Trim(Request.Form("OtherCSGMIRIssue")) != "")) {
		CSG_Issues = "<p>At this time, your application is being put on hold for the following reasons:</p><ul>";
		if (Request.Form("NoBCSAPOrHNPT") == "on") {
			CSG_Issues = CSG_Issues + "<li>There is no record of a current BCSAP application for you or we do not have a copy of your approved High-Need Part-Time (HNPT) grant form, including study period dates and signature of the Financial Aid Officer at your institution.";
		}
		if (Request.Form("BCSAPOrHNPTErrors") == "on") {
			CSG_Issues = CSG_Issues + "<li>Your current BCSAP or HNPT application contains errors and cannot be processed.";
		}
		if (Request.Form("NoFinancialNeed") == "on") {
			CSG_Issues = CSG_Issues + "<li>Your current BCSAP application does not show financial need.";
		}
		if (Request.Form("CanadaStudentLoanDefault") == "on") {
			CSG_Issues = CSG_Issues + "<li>You have a Canada Student Loan Default.  Federal policy prohibits us from providing CSG funding until the issue is resolved.";
		}
		if (Request.Form("InsufficientTimeForProcessingTSSP") == "on") {
			CSG_Issues = CSG_Issues + "<li>There is insufficient time to process your TSSP prior the end of your study period.  A completed application must be received at least 6 weeks pior to the end of your study period (if you have a BCSAP) and 4 weeks (if you have a HNPT).  You will need to reapply for a new BCSAP or HNPT."
		}
		if (Request.Form("OutstandingReceipts") == "on") {
			CSG_Issues = CSG_Issues + "<li>The Ministry does not have copies of your receipts or the unused funds from a previous CSG.  Please send the receipts and/or any unused funding to the attention of the Special Programs Unit, Ministry of Advanced Education.";
		}
		if (Request.Form("ATBCDefaultForBuyout") == "on") {
			CSG_Issues = CSG_Issues + "<li>You are in default with AT-BC.  Please contact AT-BC to resolve this status.";
		}
		if (Request.Form("ContactATBCForClarifyEquipment") == "on") {
			CSG_Issues = CSG_Issues + "<li>Contact AT-BC to clarify equipment needs.";
		}
		if (Request.Form("RequestForSecondSystem") == "on") {
			CSG_Issues = CSG_Issues + "<li>Your request for a second computer system has been denied because Ministry policy prevents us from granting you a second system within 3 years (if there are exceptional circumstances, you may submit a letter outlining those circumstances).";
		}
		if (Request.Form("IneligibleEquipment") == "on") {
			CSG_Issues = CSG_Issues + "<li>Your request for a " + (Request.Form("Comment")) + " has been denied.  This is not an allowable expense under the Federal Policy for eligible, exceptional, education-related equipment.";
		}
		if (Trim(Request.Form("OtherCSGMIRIssue")) != "") {
			CSG_Issues = CSG_Issues + "<li>" + Trim(Request.Form("OtherCSGMIRIssue")) + ".";
		}
	}

	if (CSG_Issues.length > 0) {
		CSG_Issues = CSG_Issues + "</ul><p>AT-BC will re-activate your file when the above issues have been resolved.</p>";
	}

	if ((Request.Form("CSGMissingDocumentationOne") > "0") || (Request.Form("CSGMissingDocumentationTwo") > "0") || (Request.Form("CSGMissingDocumentationThree") > "0") || (Request.Form("CSGMissingDocumentationFour") > "0") || (Request.Form("CSGMissingDocumentationFive") > "0") || (String(Request.Form("OtherCSGMissingDocumentation")) != "")) {
		CSG_Missing_Documentation = "<ul>";
		switch (String(Request.Form("CSGMissingDocumentationOne"))) {
			case "115":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Original TSSP application (faxed TSSP application forms are not acceptable).";
			break;
			case "116":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Completed TSSP application with original signatures (see highlighted sections to be filled out).";
			break;
			case "113":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of enrollment for current semester.";
			break;
			case "117":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of enrollment for following semester.";
			break;
			case "118":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of a permanent disability that meets Ministry criteria (see Ministry criteria on web site http://www.aved.gov.bc.ca/studentservices); Documentation should state if disability is permanent and include the functional impact of the disability (as described by a physician or appropriate specialist).";
			break;
			case "119":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>A letter that states that ICBC, LTD, WCB, or Department of Veterans Affairs will not or has not provided the equipment that you are requesting.";
			break;
			case "120":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Disability related rationale for requested equipment.";
			break;
		}
		switch (String(Request.Form("CSGMissingDocumentationTwo"))) {
			case "115":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Original TSSP application (faxed TSSP application forms are not acceptable).";
			break;
			case "116":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Completed TSSP application with original signatures (see highlighted sections to be filled out).";
			break;
			case "113":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of enrollment for current semester.";
			break;
			case "117":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of enrollment for following semester.";
			break;
			case "118":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of a permanent disability that meets Ministry criteria (see Ministry criteria on web site http://www.aved.gov.bc.ca/studentservices); Documentation should state if disability is permanent and include the functional impact of the disability (as described by a physician or appropriate specialist).";
			break;
			case "119":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>A letter that states that ICBC, LTD, WCB, or Department of Veterans Affairs will not or has not provided the equipment that you are requesting.";
			break;
			case "120":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Disability related rationale for requested equipment.";
			break;
		}
		switch (String(Request.Form("CSGMissingDocumentationThree"))) {
			case "115":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Original TSSP application (faxed TSSP application forms are not acceptable).";
			break;
			case "116":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Completed TSSP application with original signatures (see highlighted sections to be filled out).";
			break;
			case "113":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of enrollment for current semester.";
			break;
			case "117":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of enrollment for following semester.";
			break;
			case "118":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of a permanent disability that meets Ministry criteria (see Ministry criteria on web site http://www.aved.gov.bc.ca/studentservices); Documentation should state if disability is permanent and include the functional impact of the disability (as described by a physician or appropriate specialist).";
			break;
			case "119":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>A letter that states that ICBC, LTD, WCB, or Department of Veterans Affairs will not or has not provided the equipment that you are requesting.";
			break;
			case "120":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Disability related rationale for requested equipment.";
			break;
		}
		switch (String(Request.Form("CSGMissingDocumentationFour"))) {
			case "115":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Original TSSP application (faxed TSSP application forms are not acceptable).";
			break;
			case "116":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Completed TSSP application with original signatures (see highlighted sections to be filled out).";
			break;
			case "113":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of enrollment for current semester.";
			break;
			case "117":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of enrollment for following semester.";
			break;
			case "118":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Verification of a permanent disability that meets Ministry criteria (see Ministry criteria on web site http://www.aved.gov.bc.ca/studentservices); Documentation should state if disability is permanent and include the functional impact of the disability (as described by a physician or appropriate specialist).";
			break;
			case "119":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>A letter that states that ICBC, LTD, WCB, or Department of Veterans Affairs will not or has not provided the equipment that you are requesting.";
			break;
			case "120":
				CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>Disability related rationale for requested equipment.";
			break;
		}
		if (Trim(Request.Form("OtherCSGMissingDocumentation")) != "") {
			CSG_Missing_Documentation = CSG_Missing_Documentation + "<li>" + Trim(Request.Form("OtherCSGMissingDocumentation")) + ".";
		}
		CSG_Missing_Documentation = CSG_Missing_Documentation + "</ul>";
	}

	if (CSG_Missing_Documentation.length > 0) {
		CSG_Missing_Documentation = CSG_Missing_Documentation + "We will re-activate the file when we receive the missing documentation.";
	}

	if ((CSG_Issues.length > 0) && (CSG_Missing_Documentation.length == 0)) {
		CSG_Issues = CSG_Issues + "<p>Should you require an interim loan of equipment while you resolve the bove issues, you may contact AT-BC to apply for a short term loan of equipment.</p>";
	}

	if ((CSG_Issues.length < 0) && (CSG_Missing_Documentation.length > 0)) {
		CSG_Missing_Documentation = "<p>Your application is missing the following information and/or documentation:</p>" + CSG_Missing_Documentation
	}

	if ((CSG_Issues.length > 0) && (CSG_Missing_Documentation.length > 0)) {
		CSG_Issues = CSG_Issues + "<p>In addition to the above, your application is also missing the following information and/or documentation:</p>";
	}


	if ((Request.Form("LoanMissingDocumentationOne") > "0") || (Request.Form("LoanMissingDocumentationTwo") > "0") || (Request.Form("LoanMissingDocumentationThree") > "0") || (Request.Form("LoanMissingDocumentationFour") > "0") || (Request.Form("LoanMissingDocumentationFive") > "0") || (String(Request.Form("OtherLoanMissingDocumentation")) != "")) {
		Loan_Missing_Documentation = "<p>At this time, the application is being put on hold because the application is missing the following information and/or documentation:</p><ul>";
		switch (String(Request.Form("LoanMissingDocumentationOne"))) {
			case "121":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Complete referral form (see highlighted sections to be filled out)";
			break;
			case "122":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of enrollment for current semester; client must be enrolled in a minimum of 2 courses";
			break;
			case "123":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of enrollment for following semester; client must be enrolled in a minimum of 2 courses";
			break;
			case "125":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of a permanent disability that meets Ministry criteria; Documentation should state if disability is permanent and include the functional impact of the disability (as described by a physician or appropriate specialist)";
			break;
			case "126":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>A letter from ICBC, LTD, WCB, or DVA stating that the active claim or appeal will not result in eligibility for vocational rehabilitation services or money for vocational rehabilitation services";
			break;
			case "127":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Disability related rationale for requested equipment.";
			break;
		}
		switch (String(Request.Form("LoanMissingDocumentationTwo"))) {
			case "121":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Complete referral form (see highlighted sections to be filled out)";
			break;
			case "122":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of enrollment for current semester; client must be enrolled in a minimum of 2 courses";
			break;
			case "123":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of enrollment for following semester; client must be enrolled in a minimum of 2 courses";
			break;
			case "125":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of a permanent disability that meets Ministry criteria; Documentation should state if disability is permanent and include the functional impact of the disability (as described by a physician or appropriate specialist)";
			break;
			case "126":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>A letter from ICBC, LTD, WCB, or DVA stating that the active claim or appeal will not result in eligibility for vocational rehabilitation services or money for vocational rehabilitation services";
			break;
			case "127":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Disability related rationale for requested equipment.";
			break;
		}
		switch (String(Request.Form("LoanMissingDocumentationThree"))) {
			case "121":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Complete referral form (see highlighted sections to be filled out)";
			break;
			case "122":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of enrollment for current semester; client must be enrolled in a minimum of 2 courses";
			break;
			case "123":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of enrollment for following semester; client must be enrolled in a minimum of 2 courses";
			break;
			case "125":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of a permanent disability that meets Ministry criteria; Documentation should state if disability is permanent and include the functional impact of the disability (as described by a physician or appropriate specialist)";
			break;
			case "126":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>A letter from ICBC, LTD, WCB, or DVA stating that the active claim or appeal will not result in eligibility for vocational rehabilitation services or money for vocational rehabilitation services";
			break;
			case "127":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Disability related rationale for requested equipment.";
			break;
		}
		switch (String(Request.Form("LoanMissingDocumentationFour"))) {
			case "121":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Complete referral form (see highlighted sections to be filled out)";
			break;
			case "122":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of enrollment for current semester; client must be enrolled in a minimum of 2 courses";
			break;
			case "123":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of enrollment for following semester; client must be enrolled in a minimum of 2 courses";
			break;
			case "125":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Verification of a permanent disability that meets Ministry criteria; Documentation should state if disability is permanent and include the functional impact of the disability (as described by a physician or appropriate specialist)";
			break;
			case "126":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>A letter from ICBC, LTD, WCB, or DVA stating that the active claim or appeal will not result in eligibility for vocational rehabilitation services or money for vocational rehabilitation services";
			break;
			case "127":
				Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>Disability related rationale for requested equipment.";
			break;
		}
		if (Trim(Request.Form("OtherLoanMissingDocumentation")) != "") {
			Loan_Missing_Documentation = Loan_Missing_Documentation + "<li>" + Trim(Request.Form("OtherLoanMissingDocumentation"));
		}
		Loan_Missing_Documentation = Loan_Missing_Documentation + "</ul>";
	}

	if (Loan_Missing_Documentation.length > 0) {
		Loan_Missing_Documentation = Loan_Missing_Documentation + "We will re-activate the file when we receive the missing documentation.";
	}

	if ((Request.Form("InsufficientTimeForProcessing") == "on") || (Request.Form("ContactATBCForCIP") == "on") || (Request.Form("ATBCDefaultForLoan") == "on") || (Request.Form("EnrollmentInOneCourse") == "on") || (Request.Form("InsufficientAcademicProgress") == "on") || (Trim(Request.Form("OtherLoanMIRIssue")) != "")) {
		Loan_Issues = "<ul>";
		if (Request.Form("InsufficientTimeForProcessing") == "on") {
			var temp = "";
			if (Request.Form("Courses") == "on") {
				temp = temp + "courses, ";
			}
			if (Request.Form("JobPlacement") == "on") {
				temp = temp + "job placement, ";
			}
			if (Request.Form("JobShadow") == "on") {
				temp = temp + "job shadow, ";
			}
			if (Request.Form("PSTP") == "on") {
				temp = temp + "PSTP, ";
			}
			if (Request.Form("VolunteerPosition") == "on") {
				temp = temp + "volunteer position, ";
			}
			if (Request.Form("Practicum") == "on") {
				temp = temp + "practicum, ";
			}
			if (Trim(Request.Form("OtherComment")) != "") {
				temp = temp + Trim(Request.Form("OtherComment")) + ", ";
			}
			temp = temp.substring(0, temp.length-2);
			Loan_Issues = Loan_Issues + "<li>There is insufficient time to process the application prior to the end of the client's " + temp + ".  A complete application must be received at least 6 weeks prior to the end of the " + temp + ".";
		}
		if (Request.Form("ContactATBCForCIP") == "on") {
			Loan_Issues = Loan_Issues + "<li>Contact AT-BC to clarify equipment needs.";
		}
		if (Request.Form("ATBCDefaultForLoan") == "on") {
			Loan_Issues = Loan_Issues + "<li>" + Subject_First_Name + " is currently in default with AT-BC.  " + Subject_First_Name + " needs to contact AT-BC to arrange to return outstanding loan equipment or pay an outstanding fee.";
		}
		if (Request.Form("EnrollmentInOneCourse") == "on") {
			Loan_Issues = Loan_Issues + "<li>" + Subject_First_Name + " is enrolled in 1 course.  To be eligible for AT-BC services, the student must be enrolled in a minimum of 2 courses per semester.";
		}
		if (Request.Form("InsufficientAcademicProgress") == "on") {
			Loan_Issues = Loan_Issues + "<li>" + Subject_First_Name + " has not successfuly completed his/her courses.  To be eligible for AT-BC services, the student must demonstrate successful course completion.";
		}
		if (Trim(Request.Form("OtherLoanMIRIssue")) != "") {
			Loan_Issues = Loan_Issues + "<li>" + Trim(Request.Form("OtherLoanMIRIssue"));
		}
	}

	if (Loan_Issues.length > 0) {
		Loan_Issues = Loan_Issues + "</ul><p>AT-BC will re-activate your file when the above issues have been resolved.</p>";
	}

	if ((Loan_Issues.length > 0) && (Loan_Missing_Documentation.length == 0)) {
		Loan_Issues = "<p>At this time, the application is being put on hold for the following reasons:</p>" + Loan_Issues;
	}

	if ((Loan_Issues.length > 0) && (Loan_Missing_Documentation.length > 0)) {
		Loan_Issues = "<p>In addition to the above missing information, the following outstanding issues need to be resolved:</p>" + Loan_Issues;
	}

	Conditions_To_Maintain_Loan = "<ul>";
	if (Request.Form("EnrolledInRequiedCourses")=="on") {
		Conditions_To_Maintain_Loan = Conditions_To_Maintain_Loan + "<li>be enrolled in a minimum of " + Request.Form("NumberOfRequiredCourses") + " credit courses in either the Summer semester or Fall semester; courses must be preparing the student for employment.";
	}
	if (Request.Form("SuccessfulCompletion")=="on") {
		Conditions_To_Maintain_Loan = Conditions_To_Maintain_Loan + "<li>successfully completed courses from January to April " + CurrentYear() + ".";
	}

	if (Trim(Request.Form("OtherConditionToMaintainLoan"))!="") {
		Conditions_To_Maintain_Loan = Conditions_To_Maintain_Loan + "<li>" + Trim(Request.Form("OtherConditionToMaintainLoan"));
	}
	Conditions_To_Maintain_Loan = Conditions_To_Maintain_Loan + "</ul>";
	Reply_By_Date = Request.Form("ReplyByDate");

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
					CC_Whole_Name = CC_Whole_Name + rsContact("chvFst_Name") + " " + rsContact("chvLst_Name");
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
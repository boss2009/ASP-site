<%
var Creation_Date = "";

var Sender_Email = "";

var Recipient_Title = "";
var Recipient_Whole_Name = "";
var Recipient_First_Name = "";
var Recipient_Last_Name = "";
var Recipient_SIN = "";
var Recipient_Work_Address = "";
var Recipient_School_Address = "";
var Recipient_Job_Position = "";
var Recipient_Email = "";

var Grant_Qualification_Source = "";
var Study_Period_From = "";
var Study_Period_To = "";

var Sold_Equipment_List = "";
var Conditions = "";
var Document_Conditions = "";
var Donation = "";
var Configuration_Requested = "";
var Loan_Return = "";
var Shipping_Origin = "";
var Return_For_Donation = "";
var Return_For_Loan = "";
var Ship_For_Configuration_Requested = "";
var Not_Donation_Configuration_Loan_Return = "";


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
	Recipient_SIN = "{Recipient_SIN}";
	Recipient_Work_Address = "{Recipient_Work_Address}";
	Recipient_School_Address = "{Recipient_School_Address}";
	Recipient_Job_Position = "{Recipient_Job_Position}";
	Recipient_Email = "{Recipient_Email}";

	var Grant_Qualification_Source = "{Grant_Qualification_Source}";
	var Study_Period_From = "{Study_Period_From}";
	var Study_Period_To = "{Study_Period_To}";

	Sold_Equipment_List = "{Sold_Equipment_List}";

	Donation = "{Donation}";
	Conditions = "{Conditions}";
	Document_Conditions = "{Document_Conditions}";
	Configuration_Requested = "{Configuration_Requested}";
	Loan_Return = "{Loan_Return}";
	Shipping_Origin = "{Shipping_Origin}";
	Return_For_Donation = "{Return_For_Donation}";
	Return_For_Loan = "{Return_For_Loan}";
	Ship_For_Configuration_Requested = "{Ship_For_Configuration_Requested}";
	Not_Donation_Configuration_Loan_Return = "{Not_Donation_Configuration_Loan_Return}";
	Trainer = "{Trainer}";
	Trainer_Work_Phone = "{Trainer_Work_Phone}";
	Training_Requested = "{Training_Requested}";

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
					rsContactAddress.Source = "{call dbo.cp_Contact_Address("+ temp[1] + ",0,'','','',0,'',0,'','','',0,'','','',0,'','','','','',0,0,'Q',0)}";
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

	var RefID = 0;
	var rsBuyoutFundingSource = Server.CreateObject("ADODB.Recordset");
	rsBuyoutFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyoutFundingSource.Source = "{call dbo.cp_buyout_funding_src("+ Request.QueryString("intBuyout_req_id") + ",0,0,0,0,'Q',0)}";
	rsBuyoutFundingSource.CursorType = 0;
	rsBuyoutFundingSource.CursorLocation = 2;
	rsBuyoutFundingSource.LockType = 3;
	rsBuyoutFundingSource.Open();
	if (!rsBuyoutFundingSource.EOF)	{
		if (rsBuyoutFundingSource("bitIs_Sel_FundingSrc")=="1") RefID = rsBuyoutFundingSource.Fields.Item("intReferral_id").Value;
		rsBuyoutFundingSource.MoveNext();
	}

	if (RefID > 0) {
		var rsGrant = Server.CreateObject("ADODB.Recordset");
		rsGrant.ActiveConnection = MM_cnnASP02_STRING;
		rsGrant.Source = "{call dbo.cp_Grant_Elgbty3(0,"+ Request.QueryString("intAdult_id") + ",0,0,'Q',0)}";
		rsGrant.CursorType = 0;
		rsGrant.CursorLocation = 2;
		rsGrant.LockType = 3;
		rsGrant.Open();
		while (!rsGrant.EOF) {
			if (rsGrant.Fields.Item("intReferral_id").Value==RefID) {
				Grant_Qualification_Source = rsGrant.Fields.Item("chvGrn_Qlf_Src").Value;
				Study_Period_From = FilterDate(rsGrant.Fields.Item("dtmEligibility_start").Value);
				Study_Period_To = FilterDate(rsGrant.Fields.Item("dtmEligibility_end").Value);
			}
			rsGrant.MoveNext();
		}
	}

	var rsInventorySold = Server.CreateObject("ADODB.Recordset");
	rsInventorySold.ActiveConnection = MM_cnnASP02_STRING;
	rsInventorySold.Source = "{call dbo.cp_buyout_eqp_sold(0,"+Request.QueryString("intBuyout_req_id")+",0,0.0,'',0,0,'',0,'Q',0)}";
	rsInventorySold.CursorType = 0;
	rsInventorySold.CursorLocation = 2;
	rsInventorySold.LockType = 3;
	rsInventorySold.Open();

	var charge = 0;
	var rsChargeRate = Server.CreateObject("ADODB.Recordset");
	rsChargeRate.ActiveConnection = MM_cnnASP02_STRING;
	rsChargeRate.Source = "{call dbo.cp_charge_rate(0,'',0,0.0,0,'Q',0)}";
	rsChargeRate.CursorType = 0;
	rsChargeRate.CursorLocation = 2;
	rsChargeRate.LockType = 3;
	rsChargeRate.Open();
	while (!rsChargeRate.EOF) {
		charge += rsChargeRate.Fields.Item("fltPercentage").Value;
		rsChargeRate.MoveNext();
	}
	rsChargeRate.Close();

	if (!rsInventorySold.EOF) {
		Sold_Equipment_List = "<ul>";
	}

	var gst = 0;
	var pst = 0;
	var shipping = 0;

	var rsGST = Server.CreateObject("ADODB.Recordset");
	rsGST.ActiveConnection = MM_cnnASP02_STRING;
	rsGST.Source = "{call dbo.cp_charge_rate(1,'',0,0.0,1,'Q',0)}";
	rsGST.CursorType = 0;
	rsGST.CursorLocation = 2;
	rsGST.LockType = 3;
	rsGST.Open();
	if (!rsGST.EOF) gst = Number(rsGST.Fields.Item("fltPercentage").Value);
	rsGST.Close();

	var rsPST = Server.CreateObject("ADODB.Recordset");
	rsPST.ActiveConnection = MM_cnnASP02_STRING;
	rsPST.Source = "{call dbo.cp_charge_rate(2,'',0,0.0,1,'Q',0)}";
	rsPST.CursorType = 0;
	rsPST.CursorLocation = 2;
	rsPST.LockType = 3;
	rsPST.Open();
	if (!rsPST.EOF) pst = Number(rsPST.Fields.Item("fltPercentage").Value);
	rsPST.Close();

	var rsShipping = Server.CreateObject("ADODB.Recordset");
	rsShipping.ActiveConnection = MM_cnnASP02_STRING;
	rsShipping.Source = "{call dbo.cp_charge_rate(3,'',0,0.0,1,'Q',0)}";
	rsShipping.CursorType = 0;
	rsShipping.CursorLocation = 2;
	rsShipping.LockType = 3;
	rsShipping.Open();
	if (!rsShipping.EOF) shipping = Number(rsShipping.Fields.Item("fltPercentage").Value);
	rsShipping.Close();

	var total_sold_price = 0;
	var total_cost = 0;
	var tax = 0;
	var total_shipping = 0;

	while (!rsInventorySold.EOF) {
		if (!(rsInventorySold.Fields.Item("insEquip_Class_id").Value==null)) {
			var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
			rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
			rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW(" + rsInventorySold.Fields.Item("insEquip_Class_id").Value + ",'C',1)}";
			rsConcreteClass.CursorType = 0;
			rsConcreteClass.CursorLocation = 2;
			rsConcreteClass.LockType = 3;
			rsConcreteClass.Open();
			switch (String(rsConcreteClass.Fields.Item("chvSbjTotax").Value)) {
				//pst
				case "1":
					tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * (pst/100));
				break;
				//gst
				case "2":
					tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * (gst/100));
				break;
				//both
				case "3":
					tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * ((gst+pst)/100));
				break;
			}
		}

		Sold_Equipment_List = Sold_Equipment_List + "<li>" + Trim(rsInventorySold("chvInventory_Name"));
		total_cost += rsInventorySold.Fields.Item("fltList_Unit_Cost").Value;
		total_sold_price += rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value;
		rsInventorySold.MoveNext();
	}
	total_shipping = total_sold_price * (shipping/100);

	if (Sold_Equipment_List.length > 0) {
		Sold_Equipment_List = Sold_Equipment_List + "</ul>Total Cost of Equipment without taxes/shipping: " + FormatCurrency(total_sold_price) + "<br>";
		Sold_Equipment_List = Sold_Equipment_List + "Taxes: " + FormatCurrency(tax) + "<br>";
		Sold_Equipment_List = Sold_Equipment_List + "Shipping: " + FormatCurrency(total_shipping) + "<br>";
		Sold_Equipment_List = Sold_Equipment_List + "Total Cost of Equipment with taxes/shipping: " + FormatCurrency(total_sold_price+tax+total_shipping) + "<br>";
		Sold_Equipment_List = Sold_Equipment_List + "<br>";
	}
	rsInventorySold.Close();

	if ((Request.Form("DocumentConditionOne") > "0") || (Request.Form("DocumentConditionTwo") > "0") || (Request.Form("DocumentConditionThree") > "0") || (Request.Form("DocumentConditionFour") > "0")) {
		var rsDocumentCondition = Server.CreateObject("ADODB.Recordset");
		rsDocumentCondition.ActiveConnection = MM_cnnASP02_STRING;
		rsDocumentCondition.CursorType = 0;
		rsDocumentCondition.CursorLocation = 2;
		rsDocumentCondition.LockType = 3;

		Document_Conditions = "In order to implement the above technology plan, AT-BC needs to receive the following documentation:";
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
		Document_Conditions = Document_Conditions + "</ul>";
	}

	if (String(Trim(Request.Form("DocumentConditionOther"))) != "") {
		Conditions = "<p>Please also note that the following condition(s) applies to the technology plan:<br>" + Request.Form("DocumentConditionOther");
	}

	if (String(Trim(Request.Form("Donation"))) != "") {
		Donation = "<p>The CSG equipment will be shipped to you as soon as possible and AT-BC will then arrange to pick up the equipment to be donated to AT-BC.</p>";
		Return_For_Donation = "- Return the following equipment for donation to AT-BC:";
		Return_For_Donation = Return_For_Donation + "<ul><li>" + Trim(Request.Form("Donation")) + "</ul>";
	}

	if (String(Trim(Request.Form("ConfigurationRequested"))) != "") {
		Configuration_Requested = "<p>The configured CSG equipment will then be shipped to you as soon as possible.</p>";
		Ship_For_Configuration_Requested = "- Ship your software and/or hardware to AT-BC to enable AT-BC to configure and/or upgrade the new equipment with your existing equipment; AT-BC will contact you to arrange the pick up of the following:";
		Ship_For_Configuration_Requested = Ship_For_Configuration_Requested + "<ul><li>" + Trim(Request.Form("ConfigurationRequested")) + "</ul>";
	}

	if (String(Trim(Request.Form("LoanReturn"))) != "") {
		Loan_Return = "<p>The CSG equipment will be shipped to you as soon as possible and AT-BC will then arrange to pick up the loaned equipment to be returned to the AT-BC Loan Bank.</p>";
		Return_For_Loan = "- Return the following currently loaned equipment to the AT-BC Loan Bank:";
		Return_For_Loan = Return_For_Loan + "<ul><li>" + Trim(Request.Form("LoanReturn")) + "</ul>";
	}

	if ((String(Trim(Request.Form("Donation"))) == "") && (String(Trim(Request.Form("ConfigurationRequested"))) == "") && (String(Trim(Request.Form("LoanReturn"))) == "")){
		Not_Donation_Configuration_Loan_Return = "The CSG equipment will be shiped to you as soon as possible.";
	}

	switch (String(Request.Form("ShippingOrigin"))) {
		case "1":
			Shipping_Origin = "<p>Please note that your equipment will be shipped to you from AT-BC.</p>";
		break;
		case "2":
			Shipping_Origin = "<p>Please note that your equipment will be shipped to you directly from the vendor(s).</p>";
		break;
		case "3":
			Shipping_Origin = "<p>Please note that some of your equipment will be shipped to you from AT-BC and some will be shipped to you directly from the vendor.</p>";
		break;
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
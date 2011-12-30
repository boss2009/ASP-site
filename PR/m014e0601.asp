<!--------------------------------------------------------------------------
* File Name: m014e0601.asp
* Title: Purchase Requisition Forms & Reports
* Main SP: cp_get_purchase_requisition_02, cp_purchase_requisition_requested,
  cp_Get_Company_Address, cp_Get_Company_Address_KeyContact,
  cp_Purchase_Requisition_Note
* Description: This page retrieves values from various stored procedures and
* binds them to the pdf form variables.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="VBScript"%>
<!--#include file="../inc/VBLogin.inc"-->
<%
Dim server_address
if Left(Request.ServerVariables("remote_addr"),10) = "192.168.2." Then
	if Request.ServerVariables("LOCAL_ADDR")= "192.168.2.192" Then
		server_address = "192.168.2.192:88"
	Else
		server_address = "192.168.2.158:88"
	End if
Else
	server_address = "206.87.168.99:8080"
End if

Function FileStamp(PRID)
	DIM intLow
	DIM intHigh
	RANDOMIZE TIMER
	intLow = 100000000000000
	intHigh = 999999999999999
	FileStamp = PRID & "-" & Session("insStaff_id") & "-" & FormatNumber(Int((intHigh - intLow + 1) * RND + intHigh),0,0,0,0)
End Function

Response.Buffer = true
Dim rsPurchaseRequisition
SET rsPurchaseRequisition = Server.CreateObject("ADODB.Recordset")
rsPurchaseRequisition.ActiveConnection = MM_cnnASP02_STRING
rsPurchaseRequisition.Source = "{call dbo.cp_Get_Purchase_Requisition_02(0,0,'',1," & Request("insPurchase_Req_id") & ",0)}"
rsPurchaseRequisition.CursorType = 0
rsPurchaseRequisition.CursorLocation = 2
rsPurchaseRequisition.LockType = 3
rsPurchaseRequisition.Open()

Dim rsContractPO
SET rsContractPO = Server.CreateObject("ADODB.Recordset")
rsContractPO.ActiveConnection = MM_cnnASP02_STRING
rsContractPO.Source = "{call dbo.cp_Purchase_Requisition_Vendor(" & Request.QueryString("insPurchase_Req_id") & ",0)}"
rsContractPO.CursorType = 0
rsContractPO.CursorLocation = 2
rsContractPO.LockType = 3
rsContractPO.Open()

Dim rsInventoryRequested
Dim RequestedQuantity, RequestedDescription, RequestedLUC, RequestedTotal
Set rsInventoryRequested = Server.CreateObject("ADODB.Recordset")
rsInventoryRequested.ActiveConnection = MM_cnnASP02_STRING
rsInventoryRequested.Source = "{call dbo.cp_Purchase_Requisition_Requested(0," & Request("insPurchase_Req_id") & ",0,0,0,'',0.0,'01/01/1999',0,0,0,'Q',0)}"
rsInventoryRequested.CursorType = 0
rsInventoryRequested.CursorLocation = 2
rsInventoryRequested.LockType = 3
rsInventoryRequested.Open()

Dim rsVendor, rsVendorContact
Dim VendorAddress, ContactName, Fax, Phone
If Not rsInventoryRequested.EOF Then
	If (rsInventoryRequested.Fields.Item("insVendor_id").Value > 0) Then
		Set rsVendor = Server.CreateObject("ADODB.Recordset")
		rsVendor.ActiveConnection = MM_cnnASP02_STRING
		rsVendor.Source = "{call dbo.cp_Get_Company_Address(" & rsInventoryRequested.Fields.Item("insVendor_id").Value & ", 1)}"
		rsVendor.CursorType = 0
		rsVendor.CursorLocation = 2
		rsVendor.LockType = 3
		rsVendor.Open()
		VendorAddress = rsVendor.Fields.Item("chvAddress").Value & Chr(13) & rsVendor.Fields.Item("chvCity").Value & ", " & rsVendor.Fields.Item("chrprvst_abbv").Value & Chr(13)
		VendorAddress = VendorAddress & rsVendor.Fields.Item("chvPostal_zip").Value & Chr(13)
		Fax = ""
		Phone = ""
		If Trim(rsVendor.Fields.Item("intPhone_Type_1")) = 5 Then
			Fax = "(" & rsVendor.Fields.Item("chvPhone1_Arcd") & ") " & rsVendor.Fields.Item("chvPhone1_Num")
		End If
		If Trim(rsVendor.Fields.Item("intPhone_Type_1")) = 2 Then
			Phone = "(" &rsVendor.Fields.Item("chvPhone1_Arcd") & ") " & rsVendor.Fields.Item("chvPhone1_Num")
		End If
		If Trim(rsVendor.Fields.Item("intPhone_Type_2")) = 5 Then
			Fax = "(" & rsVendor.Fields.Item("chvPhone2_Arcd") & ") " & rsVendor.Fields.Item("chvPhone2_Num")
		End If
		If Trim(rsVendor.Fields.Item("intPhone_Type_2")) = 2 Then
			Phone = "(" & rsVendor.Fields.Item("chvPhone2_Arcd") & ") " & rsVendor.Fields.Item("chvPhone2_Num")
		End If		
		If Not Phone = "" Then
			VendorAddress = VendorAddress & "Off: " & Phone & Chr(13)
		End If
		If Not Fax = "" Then
			VendorAddress = VendorAddress & "FAX: " & Fax
		End If
		
		Set rsVendorContact = Server.CreateObject("ADODB.Recordset")
		rsVendorContact.ActiveConnection = MM_cnnASP02_STRING
		rsVendorContact.Source = "{call dbo.cp_Get_Company_Address_KeyContact(" & rsPurchaseRequisition.Fields.Item("insVendor_id").Value & ", 1, 1)}"
		rsVendorContact.CursorType = 0
		rsVendorContact.CursorLocation = 2
		rsVendorContact.LockType = 3
		rsVendorContact.Open()
		If Not rsVendorContact.EOF Then
			ContactName = rsVendorContact.Fields.Item("chvkeyContact_Fname").Value & " " & rsVendorContact.Fields.Item("chvkeyContact_Lname").Value
		Else 
			ContactName = ""
		End If
	Else
		VendorAddress = ""
		ContactName = ""
		Fax = ""
		Phone = ""	
	End If
End If

Dim RequestNotes, ReceiveNotes
Dim rsNotes
Set rsNotes = Server.CreateObject("ADODB.Recordset")
rsNotes.ActiveConnection = MM_cnnASP02_STRING
rsNotes.Source = "{call dbo.cp_Purchase_Requisition_Note(" & Request("insPurchase_Req_id") & ",'',0,0,'',0,'Q',0)}"
rsNotes.CursorType = 0
rsNotes.CursorLocation = 2
rsNotes.LockType = 3
rsNotes.Open()
While Not rsNotes.EOF
	If (rsNotes.Fields.Item("chvType_of_Note").Value = "Requested") Then
		RequestNotes = rsNotes.Fields.Item("chvNote_Desc").Value
	Else
		ReceivedNotes = rsNotes.Fields.Item("chvNote_Desc").Value
	End If
	rsNotes.MoveNext
Wend

While Not rsInventoryRequested.EOF
	RequestedQuantity = RequestedQuantity & rsInventoryRequested.Fields.Item("insPR_request_Qty_Ordered").Value & Chr(13)
	RequestedDescription = RequestedDescription & Left(rsInventoryRequested.Fields.Item("chvClass_Bundle_Name").Value,55) & "-" & rsInventoryRequested.Fields.Item("chvDescription").Value & Chr(13)
	RequestedLUC = RequestedLUC & FormatNumber(rsInventoryRequested.Fields.Item("fltPR_request_List_Unit_Cost").Value,2,0,0,-1) & Chr(13)
	RequestedTotal = RequestedTotal & FormatNumber(rsInventoryRequested.Fields.Item("fltTotal_Cost").Value,2,0,0,-1) & Chr(13)
	rsInventoryRequested.MoveNext
Wend

Dim rsInventoryReceived
Dim ReceivedQuantity, ReceivedDescription, QuantityBackOrdered, Remarks
QuantityBackOrdered = ""
Remarks = ""
Set rsInventoryReceived = Server.CreateObject("ADODB.Recordset")
rsInventoryReceived.ActiveConnection = MM_cnnASP02_STRING
rsInventoryReceived.Source = "{call dbo.cp_Purchase_Requisition_Received(" & Request("insPurchase_Req_id") & ",0,0,'',0,'',0,'Q',0)}"
rsInventoryReceived.CursorType = 0
rsInventoryReceived.CursorLocation = 2
rsInventoryReceived.LockType = 3
rsInventoryReceived.Open()
While Not rsInventoryReceived.EOF
	ReceivedQuantity = ReceivedQuantity & rsInventoryReceived.Fields.Item("intQuantity_Received").Value & Chr(13)
	ReceivedDescription = ReceivedDescription & Left(rsInventoryReceived.Fields.Item("chvClass_name").Value,70) & Chr(13)
	rsInventoryReceived.MoveNext
Wend

Dim rsStaff, StaffName
Set rsStaff = Server.CreateObject("ADODB.Recordset")
rsStaff.ActiveConnection = MM_cnnASP02_STRING
rsStaff.Source = "{call dbo.cp_Idv_Staff(" & Session("insStaff_id") & ")}"
rsStaff.CursorType = 0
rsStaff.CursorLocation = 2
rsStaff.LockType = 3
rsStaff.Open()
StaffName = rsStaff.Fields.Item("chvFst_Name").Value & " " & rsStaff.Fields.Item("chvLst_Name").Value

'	Create an instance of the Object
'
Set PRFfdf = Server.CreateObject("FdfApp.FdfApp")
Set PRRfdf = Server.CreateObject("FdfApp.FdfApp")
Set FAXfdf = Server.CreateObject("FdfApp.FdfApp")
'
' 	Use the fdfApp to feed the vars
'
Set myPRFfdf = PRFfdf.FDFCreate
Set myPRRfdf = PRRfdf.FDFCreate
Set myFAXfdf = FAXfdf.FDFCreate
'
'	Stuff the variables
'

Dim WorkOrderNumber, DateOrdered, DateReceived, PurchaseOrderNumber, OrderedBy, Vendor

If Not rsPurchaseRequisition.Fields.Item("chvWork_order").Value = vbnullstring Then
	WorkOrderNumber = rsPurchaseRequisition.Fields.Item("chvWork_order").Value
Else 
	WorkOrderNumber = ""
End If

PurchaseOrderNumber = ""
If Not rsContractPO.EOF Then
	If Not rsContractPO.Fields.Item("chvContract_PO").Value = vbnullstring Then
		PurchaseOrderNumber = rsContractPO.Fields.Item("chvContract_PO").Value
	End If
End If

If Not rsPurchaseRequisition.Fields.Item("dtsDate_Ordered").Value = vbnullstring Then
	DateOrdered = rsPurchaseRequisition.Fields.Item("dtsDate_Ordered").Value
Else 
	DateOrdered = ""
End If
If Not rsPurchaseRequisition.Fields.Item("dtsDate_Received").Value = vbnullstring Then
	DateReceived = rsPurchaseRequisition.Fields.Item("dtsDate_Received").Value
Else 
	DateReceived = ""
End If

If Not rsPurchaseRequisition.Fields.Item("chvOrdered_by").Value = vbnullstring Then
	OrderedBy = rsPurchaseRequisition.Fields.Item("chvOrdered_by").Value
Else 
	OrderedBy = ""
End If

If Not rsInventoryRequested.EOF Then
	rsInventoryRequested.MoveFirst
End If

If Not rsPurchaseRequisition.Fields.Item("insSupplier_id").Value = vbnullstring Then
	If Not rsPurchaseRequisition.Fields.Item("chvSupplier").Value = vbnullstring Then
		Vendor = rsPurchaseRequisition.Fields.Item("chvSupplier").Value
	Else
		Vendor = ""
	End If
Else 
	Vendor = ""
End If

myPRFfdf.fdfsetvalue "PurchaseRequisitionNumber", Request("insPurchase_Req_id"), false
myPRFfdf.fdfsetvalue "CompanyName", "Assistive Technology - British Columbia", false
myPRFfdf.fdfsetvalue "CompanyAddress", "Suite 112 - 1750 West 75th Avenue, Vancouver, B.C. V6P 6G2", false
myPRFfdf.fdfsetvalue "WorkOrderNumber", WorkOrderNumber, false
myPRFfdf.fdfsetvalue "DateOrdered", DateOrdered, false
myPRFfdf.fdfsetvalue "PurchaseOrderNumber", PurchaseOrderNumber, false
myPRFfdf.fdfsetvalue "Vendor", Vendor, false
myPRFfdf.fdfsetvalue "VendorAddress", VendorAddress, false
myPRFfdf.fdfsetvalue "OrderedBy", OrderedBy, false
myPRFfdf.fdfsetvalue "Notes", RequestNotes, false
myPRFfdf.fdfsetvalue "Quantity", RequestedQuantity, false
myPRFfdf.fdfsetvalue "Description", RequestedDescription, false
myPRFfdf.fdfsetvalue "ListUnitCost", RequestedLUC, false
myPRFfdf.fdfsetvalue "TotalCost", RequestedTotal, false

myPRRfdf.fdfsetvalue "PurchaseRequisitionNumber", Request("insPurchase_Req_id"), false
myPRRfdf.fdfsetvalue "CompanyName", "Assistive Technology - British Columbia", false
myPRRfdf.fdfsetvalue "CompanyAddress", "Suite 112 - 1750 West 75th Avenue, Vancouver, B.C. V6P 6G2", false
myPRRfdf.fdfsetvalue "DateReceived", DateReceived, false
myPRRfdf.fdfsetvalue "PurchaseOrderNumber", PurchaseOrderNumber, false
myPRRfdf.fdfsetvalue "Vendor", Vendor, false
myPRRfdf.fdfsetvalue "Notes", ReceiveNotes, false
myPRRfdf.fdfsetvalue "QuantityReceived", ReceivedQuantity, false
myPRRfdf.fdfsetvalue "QuantityBackOrdered", QuantityBackOrdered, false
myPRRfdf.fdfsetvalue "Description", ReceivedDescription, false
myPRRfdf.fdfsetvalue "Remarks", Remarks, false

myFAXfdf.fdfsetvalue "Date", Date(), false
myFAXfdf.fdfsetvalue "Attention", ContactName, false
myFAXfdf.fdfsetvalue "From", StaffName, false
myFAXfdf.fdfsetvalue "Fax", Fax, false
myFAXfdf.fdfsetvalue "Phone", Phone, false
myFAXfdf.fdfsetvalue "Pages", "2", false
myFAXfdf.fdfsetvalue "Message", "RE: Purchase Order #" & PurchaseOrderNumber &Chr(13)&Chr(13)&"Please proceed with the attached purchase requisition: "&Request("insPurchase_Req_id")&Chr(13)&Chr(13)&"If you have any questions, please give me a call at 269-2218."&Chr(13)&Chr(13)&"Thank you.", false

'
'	Point to your pdf file
'
myPRFfdf.fdfSetFile "http://" & server_address & "/PR/Purchase_Requisition_Form.pdf"
myPRRfdf.fdfSetFile "http://" & server_address & "/PR/Purchase_Receiving_Report.pdf"
myFAXfdf.fdfSetFile "http://" & server_address & "/PR/Facsmile_Transmission.pdf"

Response.ContentType = "text/html"
'
'	Save it to a file.  If you were going to save the actual file past the point of printing
'	You would want to create a naming convention (perhaps using social in the name)
'	Have to use the physical path so you may need to incorporate Server.mapPath in 
'	on this portion.
'

Dim filestamp1
Dim filestamp2
Dim filestamp3
filestamp1 = FileStamp(Request("insPurchase_Req_id"))
filestamp2 = FileStamp(Request("insPurchase_Req_id"))
filestamp3 = FileStamp(Request("insPurchase_Req_id"))

if Request.ServerVariables("LOCAL_ADDR")= "192.168.2.192" Then
	myPRFfdf.FDFSaveToFile "D:\WkArea\wwwroot\ASPsite\PDFTemp\Purchase_Requisition_Form_"&filestamp1&".fdf"
	myPRRfdf.FDFSaveToFile "D:\WkArea\wwwroot\ASPsite\PDFTemp\Purchase_Receiving_Report_"&filestamp2&".fdf"
	myFAXfdf.FDFSaveToFile "D:\WkArea\wwwroot\ASPsite\PDFTemp\Facsmile_Transmission_"&filestamp3&".fdf"
Else
	myPRFfdf.FDFSaveToFile "D:\Wk_area\wwwroot\ASPSite\PDFTemp\Purchase_Requisition_Form_"&filestamp1&".fdf"
	myPRRfdf.FDFSaveToFile "D:\Wk_area\wwwroot\ASPSite\PDFTemp\Purchase_Receiving_Report_"&filestamp2&".fdf"
	myFAXfdf.FDFSaveToFile "D:\Wk_area\wwwroot\ASPSite\PDFTemp\Facsmile_Transmission_"&filestamp3&".fdf"
End If
'
'	Close your Objects
'
myPRFfdf.fdfclose
myPRRfdf.fdfclose
myFAXfdf.fdfclose
set PRFfdf = nothing
set PRRfdf = nothing
set FAXfdf = nothing
set rsPurchaseRequisition = nothing
set rsNotes = nothing
set rsInventoryRequested = nothing
set rsInventoryReceived = nothing
%>
<html>
<head>
	<title>Purchase Requisition Forms & Reports</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h3>Forms & Reports</h3>
<hr>
<a href="http://<%=server_address%>/PDFTemp/Purchase_Requisition_Form_<%=filestamp1%>.fdf">Purchase Requisition Form</a><br>
<a href="http://<%=server_address%>/PDFTemp/Purchase_Receiving_Report_<%=filestamp2%>.fdf">Purchase Receiving Report</a><br>
<a href="http://<%=server_address%>/PDFTemp/Facsmile_Transmission_<%=filestamp3%>.fdf">Facsmile Transmission</a><br>
</body>
</html>
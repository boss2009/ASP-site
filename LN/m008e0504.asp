<!--------------------------------------------------------------------------
* File Name: m008e0504.asp
* Title: Loan Shipping Label
* Main SP: cp_Get_Company_Address
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

Function FileStamp(LOANID)
	DIM intLow
	DIM intHigh
	RANDOMIZE TIMER
	intLow = 100000000000000
	intHigh = 999999999999999
	FileStamp = LOANID & "-" & Session("insStaff_id") & "-" & FormatNumber(Int((intHigh - intLow + 1) * RND + intHigh),0,0,0,0)
End Function

Response.Buffer = true

Dim rsLoan 
set rsLoan = Server.CreateObject("ADODB.Recordset")
rsLoan.ActiveConnection = MM_cnnASP02_STRING
rsLoan.Source = "{call dbo.cp_get_loan_ship_name(" & Request.QueryString("intLoan_Req_id") & ",0)}"
rsLoan.CursorType = 0
rsLoan.CursorLocation = 2
rsLoan.LockType = 3
rsLoan.Open()

Dim intShip_dtl_id
intShip_dtl_id = rsLoan.Fields.Item("intShip_dtl_id").Value

Dim rsRecipient
Dim RecipientAddress
Set rsRecipient = Server.CreateObject("ADODB.Recordset")
rsRecipient.ActiveConnection = MM_cnnASP02_STRING
rsRecipient.Source = "{call dbo.cp_loan_ship_address(" & intShip_dtl_id & ",0,'','','','','',0,'','',0,'',0,'','','',0,'','','',0,'','','','','',0,'Q',0)}"
rsRecipient.CursorType = 0
rsRecipient.CursorLocation = 2
rsRecipient.LockType = 3
rsRecipient.Open()

RecipientAddress = rsRecipient.Fields.Item("chvUsr_Fstname").Value & " " & rsRecipient.Fields.Item("chvUsr_Lstname").Value & Chr(13)
RecipientAddress = RecipientAddress & rsRecipient.Fields.Item("chvAddress").Value & Chr(13)
RecipientAddress = RecipientAddress & rsRecipient.Fields.Item("chvCity").Value & ", " & rsRecipient.Fields.Item("chrprvst_abbv").Value & " " & rsRecipient.Fields.Item("chvPostal_zip").Value & Chr(13)
If rsRecipient.Fields.Item("intPhone_Type_1").Value > 0 Then
	RecipientAddress = RecipientAddress & "Tel: (" & rsRecipient.Fields.Item("chvPhone1_Arcd").Value & ") " & rsRecipient.Fields.Item("chvPhone1_Num").Value
End If

'	Create an instance of the Object
'
Set SLfdf = Server.CreateObject("FdfApp.FdfApp")
'
' 	Use the fdfApp to feed the vars
'
Set mySLfdf = SLfdf.FDFCreate
'
'	Stuff the variables
'
dim CompanyAddress
CompanyAddress = "Assistive Technology - British Columbia" & Chr(13)
CompanyAddress = CompanyAddress & "112-1750 West 75th Avenue" & Chr(13)
CompanyAddress = CompanyAddress & "Vancouver, B.C. V6P 6G2" & Chr(13)
CompanyAddress = CompanyAddress & "Tel: (604) 264-8295" & Chr(13)
CompanyAddress = CompanyAddress & "Fax: (604) 263-2267"

mySLfdf.fdfsetvalue "CompanyAddress", CompanyAddress, false
mySLfdf.fdfsetvalue "CompanyAddress2", CompanyAddress, false
mySLfdf.fdfsetvalue "RecipientAddress", RecipientAddress, false
mySLfdf.fdfsetvalue "RecipientAddress2", RecipientAddress, false
mySLfdf.fdfsetvalue "Notes", "", false

'
'	Point to your pdf file
'
mySLfdf.fdfSetFile "http://" & server_address & "/LN/Shipping_Label.pdf"

Response.ContentType = "text/html"
'
'	Save it to a file.  If you were going to save the actual file past the point of printing
'	You would want to create a naming convention (perhaps using social in the name)
'	Have to use the physical path so you may need to incorporate Server.mapPath in 
'	on this portion.
'

Dim filestamp1
filestamp1 = FileStamp(Request("intLoan_Req_id"))
if Request.ServerVariables("LOCAL_ADDR")= "192.168.2.192" Then
	mySLfdf.FDFSaveToFile "D:\Wkarea\wwwroot\ASPsite\PDFTemp\Shipping_Label_"&filestamp1&".fdf"
Else
	mySLfdf.FDFSaveToFile "D:\Wk_area\wwwroot\ASPSite\PDFTemp\Shipping_Label_"&filestamp1&".fdf"
End if
'
'	Close your Objects
'
mySLfdf.fdfclose
set SLfdf = nothing

'Response.Redirect("http://" & server_address & "/PDFTemp/Shipping_Label_" & filestamp1 & ".fdf")
%>
<html>
<head>
	<title>Loan Shipping Label</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body onLoad="window.location.href='http://<%=server_address%>/PDFTemp/Shipping_Label_<%=filestamp1%>.fdf'">
<h5>Generating Shipping Label...</h5>
If this page does not refresh within 10 seconds, contact system administrator.<br>
<br>
Possible cause of the problem:
<ul>
	<li>This computer does not have Adobe Reader installed.
	<li>Adobe Reader program error.
	<li>Server system folder has changed.
</ul></body>
</html>
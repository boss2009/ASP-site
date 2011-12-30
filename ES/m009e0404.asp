<!--------------------------------------------------------------------------
* File Name: m009e0403.asp
* Title: Equipment Service Shipping Label
* Main SP: cp_company_address
* Description: This page retrieves values from various stored procedures and
* binds them to the pdf form variables.
--------------------------------------------------------------------------->
<%@language="VBScript"%>
<!--#include file="../inc/VBLogin.inc"-->
<%
Dim server_address

' + Nov.04.2005 
'if Left(Request.ServerVariables("remote_addr"),10) = "192.168.2." Then
'	if Request.ServerVariables("LOCAL_ADDR")= "192.168.2.192" Then
'		server_address = "192.168.2.192:88"
'	Else
'		server_address = "192.168.2.158:88"
'	End if
'Else
'	server_address = "206.87.168.99:8080"
'End if
server_address = "localhost:8080/aspsite/"

Function FileStamp(BUYOUTID)
	DIM intLow
	DIM intHigh
	RANDOMIZE TIMER
	intLow = 100000000000000
	intHigh = 999999999999999
	FileStamp = BUYOUTID & "-" & Session("insStaff_id") & "-" & FormatNumber(Int((intHigh - intLow + 1) * RND + intHigh),0,0,0,0)
End Function

Response.Buffer = true

Dim RecipientAddress

RecipientAddress = Request.Form("UserName") & Chr(13)
RecipientAddress = RecipientAddress & Request.Form("StreetAddress") & Chr(13)
RecipientAddress = RecipientAddress & Request.Form("City") & ", " & Request.Form("ProvinceState") & " " & Request.Form("PostalCode") & Chr(13)
If intPhone_Type_1 > 0 Then
	RecipientAddress = RecipientAddress & "Tel: (" & Request.Form("PrimaryPhoneAreaCode") & ") " & Request.Form("PrimaryPhoneNumber")
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

CompanyAddress = "Sirius Innovations Inc" & Chr(13)
CompanyAddress = CompanyAddress & "P.O. Box 43119 Richmond Ctr PO" & Chr(13)
CompanyAddress = CompanyAddress & "Richmond, B.C. V6V 2W4" & Chr(13)
CompanyAddress = CompanyAddress & "Tel: (604) 959-8188" & Chr(13)
CompanyAddress = CompanyAddress & "Fax: (604) 959-3169"


mySLfdf.fdfsetvalue "CompanyAddress", CompanyAddress, false
mySLfdf.fdfsetvalue "CompanyAddress2", CompanyAddress, false
mySLfdf.fdfsetvalue "RecipientAddress", RecipientAddress, false
mySLfdf.fdfsetvalue "RecipientAddress2", RecipientAddress, false
mySLfdf.fdfsetvalue "Notes", "", false

'
'	Point to your pdf file
'
mySLfdf.fdfSetFile "http://" & server_address & "/ES/Shipping_Label.pdf"

Response.ContentType = "text/html"
'
'	Save it to a file.  If you were going to save the actual file past the point of printing
'	You would want to create a naming convention (perhaps using social in the name)
'	Have to use the physical path so you may need to incorporate Server.mapPath in 
'	on this portion.
'

Dim filestamp1
filestamp1 = FileStamp(Request("intEquip_Srv_id"))

' + Nov.04.2005
'if Request.ServerVariables("LOCAL_ADDR")= "192.168.2.192" Then
'	mySLfdf.FDFSaveToFile "D:\Wkarea\wwwroot\ASPsite\PDFTemp\Shipping_Label_"&filestamp1&".fdf"
'Else
'	mySLfdf.FDFSaveToFile "D:\Wk_area\wwwroot\ASPSite\PDFTemp\Shipping_Label_"&filestamp1&".fdf"
'End if	
mySLfdf.FDFSaveToFile "N:\wwwroot\ASPsite\PDFTemp\Shipping_Label_"&filestamp1&".fdf"


'
'	Close your Objects
'
mySLfdf.fdfclose
set SLfdf = nothing

'Response.Redirect("http://" & server_address & "/PDFTemp/Shipping_Label_" & filestamp1 & ".fdf")
%>
<html>
<head>
	<title>Equipment Service Shipping Label</title>
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
</ul>
</body>
</html>
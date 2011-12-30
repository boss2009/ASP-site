<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
// *** Edit Operations: declare variables

// set the form action variable
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsCategory = Server.CreateObject("ADODB.Recordset");
rsCategory.ActiveConnection = MM_cnnASP02_STRING;
rsCategory.Source = "{call dbo.cp_asp_lkup(126)}";
rsCategory.CursorType = 0;
rsCategory.CursorLocation = 2;
rsCategory.LockType = 3;
rsCategory.Open();

var rsClientTag = Server.CreateObject("ADODB.Recordset");
rsClientTag.ActiveConnection = MM_cnnASP02_STRING;
rsClientTag.Source = "{call dbo.cp_asp_lkup2(127,3,'',1,'',0)}";
rsClientTag.CursorType = 0;
rsClientTag.CursorLocation = 2;
rsClientTag.LockType = 3;
rsClientTag.Open();
var CountClientTag = 0;
while (!rsClientTag.EOF) {
	CountClientTag++;
	rsClientTag.MoveNext();
}
rsClientTag.MoveFirst();

var rsContactTag = Server.CreateObject("ADODB.Recordset");
rsContactTag.ActiveConnection = MM_cnnASP02_STRING;
rsContactTag.Source = "{call dbo.cp_asp_lkup2(127,2,'',1,'',0)}";
rsContactTag.CursorType = 0;
rsContactTag.CursorLocation = 2;
rsContactTag.LockType = 3;
rsContactTag.Open();
var CountContactTag = 0;
while (!rsContactTag.EOF) {
	CountContactTag++;
	rsContactTag.MoveNext();
}
rsContactTag.MoveFirst();

var rsASPTag = Server.CreateObject("ADODB.Recordset");
rsASPTag.ActiveConnection = MM_cnnASP02_STRING;
rsASPTag.Source = "{call dbo.cp_asp_lkup2(127,1,'',1,'',0)}";
rsASPTag.CursorType = 0;
rsASPTag.CursorLocation = 2;
rsASPTag.LockType = 3;
rsASPTag.Open();
var CountASPTag = 0;
while (!rsASPTag.EOF) {
	CountASPTag++;
	rsASPTag.MoveNext();
}
rsASPTag.MoveFirst();
%>
<html>
<head>
	<title>New Template</title>
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
			window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	var ClientTagArray = new Array(<%=CountClientTag%>);
	<%
	var x=0;
	while (!rsClientTag.EOF) {
	%>
		ClientTagArray[<%=x%>] = new Array(2);
		ClientTagArray[<%=x%>][0] = '<%=rsClientTag.Fields.Item("chvRTItm_Name").Value%>';
		ClientTagArray[<%=x%>][1] = '<%=rsClientTag.Fields.Item("chvRTItm_Code").Value%>';
	<%
		x++;
		rsClientTag.MoveNext();
	}
	%>
	
	var ContactTagArray = new Array(<%=CountContactTag%>);
	<%
	var y=0;
	while (!rsContactTag.EOF) {
	%>
		ContactTagArray[<%=y%>] = new Array(2);
		ContactTagArray[<%=y%>][0] = '<%=rsContactTag.Fields.Item("chvRTItm_Name").Value%>';
		ContactTagArray[<%=y%>][1] = '<%=rsContactTag.Fields.Item("chvRTItm_Code").Value%>';
	<%
		y++;
		rsContactTag.MoveNext();
	}
	%>
	
	var ASPTagArray = new Array(<%=CountASPTag%>);
	<%
	var z=0;
	while (!rsASPTag.EOF) {
	%>
		ASPTagArray[<%=z%>] = new Array(2);
		ASPTagArray[<%=z%>][0] = '<%=rsASPTag.Fields.Item("chvRTItm_Name").Value%>';
		ASPTagArray[<%=z%>][1] = '<%=rsASPTag.Fields.Item("chvRTItm_Code").Value%>';
	<%
		z++;
		rsASPTag.MoveNext();
	}
	%>
		
	function Save(){
		document.frm0604A.submit();
	}
	
	function addOption(txt, val, obj){
		var oOption=document.createElement("OPTION");
		oOption.text = txt;
		oOption.value = val;
		obj.add(oOption);
	}
	
	function ChangeHeaderTagCategory(category){
	  	while (document.frmTest.HeaderTag.length > 0){
    			document.frmTest.HeaderTag.remove(0);
		}
		switch(category){
			//ASP
			case "1":
				for (i = 0; i < <%=CountASPTag%>; i++) {
					addOption(ASPTagArray[i][0],ASPTagArray[i][1], document.frmTest.HeaderTag);
				}
			break;
			//Client
			case "2":
				for (i = 0; i < <%=CountClientTag%>; i++) {
					addOption(ClientTagArray[i][0],ClientTagArray[i][1], document.frmTest.HeaderTag);
				}
			break;
			//Contact
			case "3":
				for (i = 0; i < <%=CountContactTag%>; i++) {
					addOption(ContactTagArray[i][0],ContactTagArray[i][1], document.frmTest.HeaderTag);
				}
			break;
		}	
	}

	function ChangeFooterTagCategory(category){
	  	while (document.frmTest.FooterTag.length > 0){
    			document.frmTest.FooterTag.remove(0);
		}
		switch(category){
			//ASP
			case "1":
				for (i = 0; i < <%=CountASPTag%>; i++) {
					addOption(ASPTagArray[i][0],ASPTagArray[i][1], document.frmTest.FooterTag);
				}
			break;
			//Client
			case "2":
				for (i = 0; i < <%=CountClientTag%>; i++) {
					addOption(ClientTagArray[i][0],ClientTagArray[i][1], document.frmTest.FooterTag);
				}
			break;
			//Contact
			case "3":
				for (i = 0; i < <%=CountContactTag%>; i++) {
					addOption(ContactTagArray[i][0],ContactTagArray[i][1], document.frmTest.FooterTag);
				}
			break;
		}	
	}
	
	function InsertHeaderTag(){
		document.frmTest.HeaderBlock.value = document.frmTest.HeaderBlock.value + document.frmTest.HeaderTag.value;
	}

	function InsertFooterTag(){
		document.frmTest.FooterBlock.value = document.frmTest.FooterBlock.value + document.frmTest.FooterTag.value;
	}
	
	function Init(){
		ChangeHeaderTagCategory(document.frmTest.HeaderTagCategory.value);
		ChangeFooterTagCategory(document.frmTest.FooterTagCategory.value);		
	}
	</script>
</head>
<body onLoad="Init();" >
<form name="frmTest" method="POST" action="<%=MM_editAction%>">
<h5>Header</h5>
<select name="HeaderTagCategory" onChange="ChangeHeaderTagCategory(this.value);" tabindex="1" accesskey="F">
	<option value="2">Client
	<option value="3">Contact
	<option value="1">ASP
</select>
<select name="HeaderTag" tabindex="2">
</select>
<input type="button" style="btnstyle" tabindex="3" value="Insert Tag" onClick="InsertHeaderTag();" class="btnstyle">
<br>
<br>
<textarea name="HeaderBlock" rows="10" cols="90" tabindex="4"></textarea>
<hr>
<h5>Body</h5>
<br>
<br>
<textarea name="BodyBlock" rows="10" cols="90" tabinde=""></textarea>
<hr>
<h5>Footer</h5>
<select name="FooterTagCategory" onChange="ChangeFooterTagCategory(this.value);" tabindex="5">
	<option value="2">Client
	<option value="3">Contact
	<option value="1">ASP
</select>
<select name="FooterTag" tabindex="6">
</select>
<input type="button" style="btnstyle" tabindex="7" value="Insert Tag" onClick="InsertFooterTag();" class="btnstyle">
<br>
<br>
<textarea name="FooterBlock" rows="10" cols="90" tabindex="8"></textarea>
<hr>
</form>
</body>
</html>
<%
rsContactTag.Close();
rsASPTag.Close();
rsClientTag.Close();
%>
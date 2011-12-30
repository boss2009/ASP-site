<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}


var MM_abortEdit = false;
var MM_editQuery = "";

if (String(Request("MM_update")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable  = "dbo.tbl_crsp_hstry";
  var MM_editColumn = "intLetter_id";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "UpdateSuccessful.asp?page=m001q0901.asp&intAdult_id="+Request.QueryString("intAdult_id");
  var MM_fieldsStr = "Subject|value|Template|value|Type|value|DocumentName|value|SentDate|value";
  var MM_columnsStr = "intSubject_id|none,none,NULL|insTemplate_id|none,none,NULL|intRecipient_type|none,none,NULL|chvLetter_Name|',none,''|dtsSend_date|',none,''";

  // create the MM_fields and MM_columns arrays
  var MM_fields = MM_fieldsStr.split("|");
  var MM_columns = MM_columnsStr.split("|");
  
  // set the form values
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    MM_fields[i+1] = String(Request.Form(MM_fields[i]));
  }

  // append the query string to the redirect URL
  if (MM_editRedirectUrl && Request.QueryString && Request.QueryString.length > 0) {
    MM_editRedirectUrl += ((MM_editRedirectUrl.indexOf('?') == -1)?"?":"&") + Request.QueryString;
  }
}
%>
<%
// *** Update Record: construct a sql update statement and execute it

if (String(Request("MM_update")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  // create the sql update statement
  MM_editQuery = "update " + MM_editTable + " set ";
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    var formVal = MM_fields[i+1];
    var MM_typesArray = MM_columns[i+1].split(",");
    var delim =    (MM_typesArray[0] != "none") ? MM_typesArray[0] : "";
    var altVal =   (MM_typesArray[1] != "none") ? MM_typesArray[1] : "";
    var emptyVal = (MM_typesArray[2] != "none") ? MM_typesArray[2] : "";
    if (formVal == "" || formVal == "undefined") {
      formVal = emptyVal;
    } else {
      if (altVal != "") {
        formVal = altVal;
      } else if (delim == "'") { // escape quotes
        formVal = "'" + formVal.replace(/'/g,"''") + "'";
      } else {
        formVal = delim + formVal + delim;
      }
    }
    MM_editQuery += ((i != 0) ? "," : "") + MM_columns[i] + " = " + formVal;
  }
  MM_editQuery += " where " + MM_editColumn + " = " + MM_recordId;

  if (!MM_abortEdit) {
    // execute the update
    var MM_editCmd = Server.CreateObject('ADODB.Command');
    MM_editCmd.ActiveConnection = MM_editConnection;
    MM_editCmd.CommandText = MM_editQuery;
    MM_editCmd.Execute();
    MM_editCmd.ActiveConnection.Close();

    if (MM_editRedirectUrl) {
      Response.Redirect(MM_editRedirectUrl);
    }
  }
}

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_ClnCtact("+ Request.QueryString("intAdult_id") + ")}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

var rsTemplate = Server.CreateObject("ADODB.Recordset");
rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
rsTemplate.Source = "{call dbo.cp_Letter_template(0,1,'',0,'',0,0,0,0,0,0,2,'Q',0)}";
rsTemplate.CursorType = 0;
rsTemplate.CursorLocation = 2;
rsTemplate.LockType = 3;
rsTemplate.Open();
var count = 0;
while (!rsTemplate.EOF) {
	count++;
	rsTemplate.MoveNext();
}
rsTemplate.MoveFirst();

var rsCorrespondence = Server.CreateObject("ADODB.Recordset");
rsCorrespondence.ActiveConnection = MM_cnnASP02_STRING;
rsCorrespondence.Source = "{call dbo.cp_Idv_Crsp_hstry("+ Request.QueryString("intLetter_id") + ")}";
rsCorrespondence.CursorType = 0;
rsCorrespondence.CursorLocation = 2;
rsCorrespondence.LockType = 3;
rsCorrespondence.Open();
%>
<html>
<head>
	<title>Update Correspondence</title>
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
			case 85:
				//alert("U");
				document.frm0901.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	var DocumentArray = new Array(<%=count%>);
	<% 
	var i = 0;
	while (!rsTemplate.EOF) {
	%>
		DocumentArray[<%=i%>] = new Array(3);
		DocumentArray[<%=i%>][0] = <%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>;
		DocumentArray[<%=i%>][1] = "<%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%>";
		DocumentArray[<%=i%>][2] = "<%=(rsTemplate.Fields.Item("chvFileName").Value)%>";
	<%
		rsTemplate.MoveNext();
		i++;
	}
	rsTemplate.MoveFirst();
	%>
	function Save(){
		if (!CheckDate(document.frm0901.SentDate.value)){
			alert("Invalid Sent Date.");
			document.frm0901.SentDate.focus();
			return ;
		}
		if (Trim(document.frm0901.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0901.DocumentName.focus();
			return ;
		}
		document.frm0901.submit();
	}
	
	function GenerateLetter(){
		var FileName = DocumentArray[document.frm0901.Template.selectedIndex][2];
		var ClientID = <%=Request.QueryString("intAdult_id")%>;
		var ContactID = document.frm0901.Contact.value; 
		var WinWord = window.open("../TPL/"+FileName+"?intAdult_id="+ClientID+"&intContact_id="+ContactID);
	}
	</script>
</head>
<body onLoad="document.frm0901.Subject.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0901">
<h5>Update Correspondence</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Subject:</td>
		<td nowrap><select name="Subject" tabindex="1" accesskey="F">
		<% 
		while (!rsClient.EOF) {
		%>
			<option value="<%=(rsClient.Fields.Item("intAdult_Id").Value)%>" <%=((rsClient.Fields.Item("intAdult_Id").Value == rsCorrespondence.Fields.Item("intSubject_id").Value)?"SELECTED":"")%> ><%=(rsClient.Fields.Item("chvName").Value)%></option>
		<%
			rsClient.MoveNext();
		}
		%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Template:</td>
		<td nowrap><select name="Template" tabindex="2">
		<% 
		while (!rsTemplate.EOF) {
		%>
			<option value="<%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>" <%=((rsTemplate.Fields.Item("insTemplate_id").Value == rsCorrespondence.Fields.Item("insTemplate_id").Value)?"SELECTED":"")%> ><%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%></option>
		<%
		rsTemplate.MoveNext();
		}
		%>
        </select></td>
	</tr>
	<tr>
		<td nowrap>Contact:</td>
		<td nowrap><select name="Contact" tabindex="3">
		<% 
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=((rsContact.Fields.Item("intContact_id").Value == 1)?"SELECTED":"")%> ><%=rsContact.Fields.Item("chvName").Value%>, <%=(rsContact.Fields.Item("chvRelationship").Value)%></option>
		<%
			rsContact.MoveNext();
		}
		%>		
		</select></td>
	</tr>	
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="4">
			<option value="4" <%=((rsCorrespondence.Fields.Item("intRecipient_type").Value == 4)?"SELECTED":"")%>>Letter
			<option value="0" <%=((rsCorrespondence.Fields.Item("intRecipient_type").Value == 0)?"SELECTED":"")%>>Form
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Document Name:</td>
		<td nowrap><input type="text" name="DocumentName" value="<%=(rsCorrespondence.Fields.Item("chvLetter_Name").Value)%>" maxlength="50" size="30" tabindex="5"></td>
    </tr>
    <tr> 
		<td nowrap>Sent Date:</td>
		<td nowrap>
			<input type="text" name="SentDate" value="<%=FilterDate(rsCorrespondence.Fields.Item("dtsSend_date").Value)%>" size="11" maxlength="10" tabindex="6" accesskey="L" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="7" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="8" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="9" onClick="history.back();" class="btnstyle"></td>
		<td><input type="button" value="Generate" onClick="GenerateLetter();" tabindex="10" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsCorrespondence.Fields.Item("intLetter_id").Value %>">
</form>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
rsTemplate.Close();
rsCorrespondence.Close();
%>
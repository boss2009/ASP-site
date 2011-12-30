<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

// boolean to abort record edit
var MM_abortEdit = false;

// query string to execute
var MM_editQuery = "";
%>
<%
// *** Insert Record: set variables

if (String(Request("MM_insert")) != "undefined") {

  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable  = "dbo.tbl_crsp_hstry";
  var MM_editRedirectUrl = "InsertSuccessful.html";
  var MM_fieldsStr = "Subject|value|Template|value|Type|value|DocumentName|value|DateSent|value";
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
// *** Insert Record: construct a sql insert statement and execute it

if (String(Request("MM_insert")) != "undefined") {

  // create the sql insert statement
  var MM_tableValues = "", MM_dbValues = "";
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
    MM_tableValues += ((i != 0) ? "," : "") + MM_columns[i];
    MM_dbValues += ((i != 0) ? "," : "") + formVal;
  }
  MM_editQuery = "insert into " + MM_editTable + " (" + MM_tableValues + ") values (" + MM_dbValues + ")";

  if (!MM_abortEdit) {
    // execute the insert
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
%>
<%
var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_ClnCtact("+ Request.QueryString("intAdult_id") + ")}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client(" + Request.QueryString("intAdult_id") + ")}";
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
%>
<html>
<head>
	<title>New Correspondence</title>
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
	function Save(){
		if (!CheckDate(document.frm0901.DateSent.value)){
			alert("Invalid Date Sent.");
			document.frm0901.DateSent.focus();
			return ;
		}
		if (Trim(document.frm0901.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0901.DocumentName.focus();
			return ;
		}
		document.frm0901.submit();
	}
	</script>
</head>
<body onLoad="document.frm0901.Subject.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0901">
<h5>New Correspondence</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Subject:</td>
		<td nowrap><select name="Subject" tabindex="1" accesskey="F">
		<% 
		while (!rsClient.EOF) {
		%>
			<option value="<%=(rsClient.Fields.Item("intAdult_Id").Value)%>" <%=((rsClient.Fields.Item("intAdult_Id").Value == "Request.QueryString(\"intAdult_id\")")?"SELECTED":"")%> ><%=(rsClient.Fields.Item("chvName").Value)%></option>
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
			<option value="<%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>" <%=((rsTemplate.Fields.Item("insTemplate_id").Value == 1)?"SELECTED":"")%> ><%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%></option>
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
			<option value="4" SELECTED>Letter
			<option value="0">Form
		</select></td> 
	</tr>
    <tr> 
		<td nowrap>Document Name:</td>
		<td nowrap><input type="text" name="DocumentName" maxlength="50" size="30" tabindex="5"></td>
    </tr>
    <tr> 
		<td nowrap>Date Sent:</td>
		<td nowrap>
			<input type="text" name="DateSent" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="6" accesskey="L" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="7" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="8" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
rsTemplate.Close();
%>
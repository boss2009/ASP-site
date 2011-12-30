<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
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
  var MM_editTable  = "dbo.tbl_staff";
  var MM_editRedirectUrl = "AddDeleteSuccessful.asp?action=Add";
  var MM_fieldsStr = "Title|value|FirstName|value|LastName|value|Region|value|JobTitle|value|Notes|value|CreatedBy|value|CreatedOn|value|IsActive|value";
  var MM_columnsStr = "insTitle_Typ_id|none,none,NULL|chvFst_Name|',none,''|chvLst_Name|',none,''|insRegion_Num|none,none,NULL|chvJobTitle|',none,''|chvNotes|',none,''|insCreator_user_id|none,none,NULL|dtsRec_Create_Date|',none,''|bitis_active|none,1,0";

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

var rsTitleType = Server.CreateObject("ADODB.Recordset");
rsTitleType.ActiveConnection = MM_cnnASP02_STRING;
rsTitleType.Source = "{call dbo.cp_TITLE_type(0,0)}";
rsTitleType.CursorType = 0;
rsTitleType.CursorLocation = 2;
rsTitleType.LockType = 3;
rsTitleType.Open();

var rsRegion = Server.CreateObject("ADODB.Recordset");
rsRegion.ActiveConnection = MM_cnnASP02_STRING;
rsRegion.Source = "{call dbo.cp_AC_Region(0,1,0)}";
rsRegion.CursorType = 0;
rsRegion.CursorLocation = 2;
rsRegion.LockType = 3;
rsRegion.Open();
%>
<html>
<head>
	<title>New Staff</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../../js/MyFunctions.js"></script>
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
		if (!CheckTextArea(document.frm0310.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		if (Trim(document.frm0310.LastName.value)=="") {
			alert("Enter Last Name.");
			document.frm0310.LastName.focus();
			return ;
		}
		document.frm0310.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0310.Title.focus();">
<form name="frm0310" method="POST" action="<%=MM_editAction%>">
<h5>New Staff</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Title:</td>
		<td><select name="Title" accesskey="F" tabindex="1">
			<% 
			while (!rsTitleType.EOF) {
			%>
				<option value="<%=(rsTitleType.Fields.Item("insTitle_Typ_id").Value)%>"><%=(rsTitleType.Fields.Item("chvtitle").Value)%></option>
			<%
				rsTitleType.MoveNext();
			}
			%>
        </select></td>
	</tr>
	<tr>
		<td nowrap>First Name:</td>
		<td><input type="text" name="FirstName" maxlength="40" size="40" tabindex="2"></td>
    </tr>
	<tr>
		<td nowrap>Last Name:</td>
		<td><input type="text" name="LastName" maxlength="40" size="40" tabindex="3"></td>
    </tr>
	<tr>
		<td nowrap>Is Active:</td> 
		<td><input type="checkbox" name="IsActive" value="1" tabindex="4" class="chkstyle"></td>
    </tr>
	<tr>
		<td nowrap>Region:</td>
		<td><select name="Region" tabindex="5">
			<% 
			while (!rsRegion.EOF) {
			%>
				<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>"><%=(rsRegion.Fields.Item("chvname").Value)%></option>
			<%
				rsRegion.MoveNext();
			}
			%>
        </select></td>
	</tr>	
    <tr> 
		<td>Job Title:</td>
		<td><input type="text" name="JobTitle" tabindex="6"></td>
	</tr>
	<tr>
		<td nowrap>Modified By:</td>
		<td><input type="text" name="CreatedBy" value="<%=Session("insStaff_id")%>" readonly tabindex="7" size="5" maxlength=5 ></td>
	</tr>
	<tr>
		<td nowrap>Modified On:</td>
		<td>
			<input type="text" name="CreatedOn" value="<%=CurrentDate()%>" readonly tabindex="8" size="11" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>	
	<tr>
	 	<td valign="top">Notes:</td>
		<td><textarea name="Notes" cols="65" rows="3" tabindex="9" accesskey="L"></textarea></td>		
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="10" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="11" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
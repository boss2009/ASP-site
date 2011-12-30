<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
function CurrentDate(){
   var d, s = "";
   d = new Date();
   s += d.getYear();
   return(s);
}

var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var MM_abortEdit = false;
var MM_editQuery = "";

if (String(Request("MM_insert")) == "true") {
  var MM_editConnection = MM_cnnASP02_STRING;
  var MM_editTable  = "dbo.tbl_Enrollment";
  var MM_editRedirectUrl = "InsertSuccessful.html";
  var MM_fieldsStr = "ReferralDate|value|Semester|value|Year|value|NumberOfCourses|value|CourseType|value|EligibleForASP|value|Comments|value";
  var MM_columnsStr = "intReferral_id|none,none,NULL|insSmstr_id|none,none,NULL|insYear|none,none,NULL|insNum_of_Course|none,none,NULL|insCourse_id|none,none,NULL|bitIsElgb4_ASP|none,none,NULL|chvComment|',none,''";

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

var rsSemester = Server.CreateObject("ADODB.Recordset");
rsSemester.ActiveConnection = MM_cnnASP02_STRING;
rsSemester.Source = "{call dbo.cp_Semester}";
rsSemester.CursorType = 0;
rsSemester.CursorLocation = 2;
rsSemester.LockType = 3;
rsSemester.Open();

var rsCourseType = Server.CreateObject("ADODB.Recordset");
rsCourseType.ActiveConnection = MM_cnnASP02_STRING;
rsCourseType.Source = "{call dbo.cp_CseType}";
rsCourseType.CursorType = 0;
rsCourseType.CursorLocation = 2;
rsCourseType.LockType = 3;
rsCourseType.Open();

var rsReferral = Server.CreateObject("ADODB.Recordset");
rsReferral.ActiveConnection = MM_cnnASP02_STRING;
rsReferral.Source = "{call dbo.cp_Referrals("+ Request.QueryString("intAdult_id") + ")}";
rsReferral.CursorType = 0;
rsReferral.CursorLocation = 2;
rsReferral.LockType = 3;
rsReferral.Open();
%>
<html>
<head>
	<title>New Education Documentation</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				document.frm0602.submit();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>	
</head>
<body onLoad="javascript:document.frm0602.ReferralDate.focus()" >
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0602">
<h5>New Education Documentation</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Referral Date:</td>
		<td nowrap><select name="ReferralDate" tabindex="1" accesskey="F">
			<% 
			while (!rsReferral.EOF) {
			%>
				<option value="<%=(rsReferral.Fields.Item("intReferral_id").Value)%>" <%=((rsReferral.Fields.Item("intReferral_id").Value == 0)?"SELECTED":"")%> ><%=(rsReferral.Fields.Item("dtsRefral_date").Value)%></option>
			<%
				rsReferral.MoveNext();
			}
			%>
		</select></td>
	</tr>
    <tr> 
		<td nowrap>Semester:</td>
		<td nowrap><select name="Semester" tabindex="2">
		<% 
		while (!rsSemester.EOF) {
		%>
			<option value="<%=(rsSemester.Fields.Item("insSmstr_id").Value)%>" <%=((rsSemester.Fields.Item("insSmstr_id").Value == 0)?"SELECTED":"")%> ><%=(rsSemester.Fields.Item("chvsmstr_name").Value)%></option>
		<%
			rsSemester.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>Year:</td>
		<td nowrap><input type="text" name="Year" value="<%=CurrentDate()%>" tabindex="3" size="4" maxlength="4" onKeypress="AllowNumericOnly();" ></td>		
	</tr>
    <tr> 
		<td nowrap># of Courses:</td>
		<td nowrap><select name="NumberOfCourses" tabindex="4">
		<%
		for (var i=1; i < 15; i++){
		%>
			<option value="<%=i%>" <%=((i == 1)?"SELECTED":"")%>><%=i%>
		<%
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>Course Type:</td>
		<td nowrap><select name="CourseType" tabindex="5">
		<% 
		while (!rsCourseType.EOF) {
		%>
			<option value="<%=(rsCourseType.Fields.Item("insCourse_id").Value)%>" <%=((rsCourseType.Fields.Item("insCourse_id").Value == 0)?"SELECTED":"")%> ><%=(rsCourseType.Fields.Item("chvcourse_name").Value)%></option>
		<%
			rsCourseType.MoveNext();
		}
		%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Eligible for ASP:</td>
		<td nowrap><select name="EligibleForASP" tabindex="6">
			<option value="1">Yes
			<option value="0" SELECTED>No
		</select></td>		
    </tr>
    <tr> 
		<td nowrap valign="top">Comment:</td>
		<td nowrap valign="top"><textarea name="Comments" rows="3" cols="64" tabindex="7" accesskey="L"></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="submit" value="Save" tabindex="8" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="9" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsSemester.Close();
rsCourseType.Close();
rsReferral.Close();
%>
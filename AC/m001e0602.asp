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
  var MM_editTable  = "dbo.tbl_Enrollment";
  var MM_editColumn = "intEnroll_id";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "UpdateSuccessful.asp?page=m001q0602.asp&intAdult_id="+Request.QueryString("intAdult_id");
  var MM_fieldsStr = "Semester|value|Year|value|NumberOfCourses|value|CourseType|value|EligibleForASP|value|Comments|value";
  var MM_columnsStr = "insSmstr_id|none,none,NULL|insYear|none,none,NULL|insNum_of_Course|none,none,NULL|insCourse_id|none,none,NULL|bitIsElgb4_ASP|none,none,NULL|chvComment|',none,''";

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

var rsSemester = Server.CreateObject("ADODB.Recordset");
rsSemester.ActiveConnection = MM_cnnASP02_STRING;
rsSemester.Source = "{call dbo.cp_Semester}";
rsSemester.CursorType = 0;
rsSemester.CursorLocation = 2;
rsSemester.LockType = 3;
rsSemester.Open();

var rsDocumentation = Server.CreateObject("ADODB.Recordset");
rsDocumentation.ActiveConnection = MM_cnnASP02_STRING;
rsDocumentation.Source = "{call dbo.cp_Idv_Edu_Doc("+ Request.QueryString("intEnroll_id") + ")}";
rsDocumentation.CursorType = 0;
rsDocumentation.CursorLocation = 2;
rsDocumentation.LockType = 3;
rsDocumentation.Open();

var rsCaseType = Server.CreateObject("ADODB.Recordset");
rsCaseType.ActiveConnection = MM_cnnASP02_STRING;
rsCaseType.Source = "{call dbo.cp_CseType}";
rsCaseType.CursorType = 0;
rsCaseType.CursorLocation = 2;
rsCaseType.LockType = 3;
rsCaseType.Open();

var rsEnrollmentReferral = Server.CreateObject("ADODB.Recordset");
rsEnrollmentReferral.ActiveConnection = MM_cnnASP02_STRING;
rsEnrollmentReferral.Source = "{call dbo.cp_Idv_EnrollRefral("+ Request.QueryString("intEnroll_id") + ")}";
rsEnrollmentReferral.CursorType = 0;
rsEnrollmentReferral.CursorLocation = 2;
rsEnrollmentReferral.LockType = 3;
rsEnrollmentReferral.Open();
%>
<html>
<head>
	<title>Update Education Documentation</title>
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
			case 85 :
				//alert("U");
				document.frm0602.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
</head>
<body onLoad="javascript:document.frm0602.Semester.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0602">
<h5>Update Education Documentation</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Semester:</td>
		<td nowrap><select name="Semester" tabindex="1" accesskey="F">
		<% 
		while (!rsSemester.EOF) {
		%>
			<option value="<%=(rsSemester.Fields.Item("insSmstr_id").Value)%>" <%=((rsSemester.Fields.Item("insSmstr_id").Value == rsDocumentation.Fields.Item("insSmstr_id").Value)?"SELECTED":"")%>><%=(rsSemester.Fields.Item("chvsmstr_name").Value)%></option>
		<%
			rsSemester.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>Year:</td>
		<td nowrap><input type="text" name="Year" size="4" maxlength="4" value="<%=rsDocumentation.Fields.Item("insYear").Value%>" tabindex="2" onKeypress="AllowNumericOnly();"></td>
    </tr>
    <tr> 
		<td nowrap># of Courses:</td>
		<td nowrap><select name="NumberOfCourses" tabindex="3">
		<%
		for (var i=1; i < 10; i++){
		%>
			<option value="<%=i%>" <%=((i==rsDocumentation.Fields.Item("insNum_of_Course").Value)?"SELECTED":"")%>><%=i%>
		<%
		}
		%>		
		</select></td>
	</tr>
	<tr>
		<td nowrap>Course Type:</td>
		<td nowrap><select name="CourseType" tabindex="4">
		<% 
		while (!rsCaseType.EOF) {
		%>
			<option value="<%=(rsCaseType.Fields.Item("insCourse_id").Value)%>" <%=((rsCaseType.Fields.Item("insCourse_id").Value == rsDocumentation.Fields.Item("insCourse_id").Value)?"SELECTED":"")%> ><%=(rsCaseType.Fields.Item("chvcourse_name").Value)%></option>
		<%
			rsCaseType.MoveNext();
		}
		%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Eligible for ASP:</td>
		<td nowrap><select name="EligibleForASP" tabindex="5">
			<option value="1" <%=((rsDocumentation.Fields.Item("bitIsElgb4_ASP").Value == 1)?"SELECTED":"")%>>Yes
			<option value="0" <%=((rsDocumentation.Fields.Item("bitIsElgb4_ASP").Value == 0)?"SELECTED":"")%>>No
		</select></td>
	</tr>
    <tr> 
		<td nowrap valign="top">Comments:</td>
		<td nowrap valign="top"><textarea name="Comments" rows="3" cols="65" tabindex="6" accesskey="L"><%=(rsDocumentation.Fields.Item("chvComment").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="submit" value="Save" tabindex="7" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="8" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="9" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsEnrollmentReferral.Fields.Item("intEnroll_id").Value %>">
</form>
</body>
</html>
<%
rsSemester.Close();
rsDocumentation.Close();
rsCaseType.Close();
rsEnrollmentReferral.Close();
%>
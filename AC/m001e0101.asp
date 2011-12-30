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
  var MM_editTable  = "dbo.tbl_Adult_Client";
  var MM_editColumn = "intAdult_id";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "UpdateSuccessful.asp?page=m001e0101.asp&intAdult_id="+Request.QueryString("intAdult_id");
  var MM_fieldsStr = "FirstName|value|Pen|value|MiddleName|value|Gender|value|LastName|value|Sin|value|DateOfBirth|value|Region|value|SetBCServed|value|PRCVIServed|value|Status|value|CaseManager|value|PrimaryDisability|value|IsFirstNation|value|SecondaryDisability|value|ProgramStanding|value";
  var MM_columnsStr = "chvFst_name|',none,''|chrPEN_num|',none,''|chvMdl_name|',none,''|bitGender_is_male|none,none,NULL|chvLst_name|',none,''|chrSIN_no|',none,''|dtsBirth_date|',none,''|insRegion_num|none,none,NULL|bitIs_Prx_SETBC|none,1,0|bitIs_Prx_PRCVI|none,1,0|insStdnt_Status_id|none,none,NULL|insCase_mngr_id|none,none,NULL|insDsbty1_id|none,none,NULL|bitIs_FirstNations|none,none,NULL|insDsbty2_id|none,none,NULL|bitIsDefault_asp|none,none,NULL";

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

var rsRegion = Server.CreateObject("ADODB.Recordset");
rsRegion.ActiveConnection = MM_cnnASP02_STRING;
rsRegion.Source = "{call dbo.cp_Region}";
rsRegion.CursorType = 0;
rsRegion.CursorLocation = 2;
rsRegion.LockType = 3;
rsRegion.Open();

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_StdStatus}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();

var rsDisability = Server.CreateObject("ADODB.Recordset");
rsDisability.ActiveConnection = MM_cnnASP02_STRING;
rsDisability.Source = "{call dbo.cp_StdDsbty}";
rsDisability.CursorType = 0;
rsDisability.CursorLocation = 2;
rsDisability.LockType = 3;
rsDisability.Open();

var rsCaseManager = Server.CreateObject("ADODB.Recordset");
rsCaseManager.ActiveConnection = MM_cnnASP02_STRING;
rsCaseManager.Source = "{call dbo.cp_CaseMgr}";
rsCaseManager.CursorType = 0;
rsCaseManager.CursorLocation = 2;
rsCaseManager.LockType = 3;
rsCaseManager.Open();

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
%>									
<html>
<head>
	<title>General Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83:
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0101.reset();
			break;
		}
	}
	</script>	
	<script language="Javascript">	
	function Save(){
		if (!CheckSIN(document.frm0101.Sin.value)) {
			alert("Invalid Social Insurance Number.");
			document.frm0101.Sin.focus();
			return ;
		}	
		if (!CheckDate(document.frm0101.DateOfBirth.value)){
			alert("Invalid Date of Birth.");
			document.frm0101.DateOfBirth.focus();
			return ;
		}
		document.frm0101.Sin.value = LeaveDigits(document.frm0101.Sin.value);
		document.frm0101.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0101.FirstName.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0101">
<h5>General Information</h5>
<hr>
<table cellpadding="1" cellspacing="1" width="100%">
	<tr> 
		<td nowrap>ASP ID:</td>
		<td nowrap width="150"><input type="text" name="ASPID" value="<%=Request.QueryString("intAdult_id")%>" readonly style="border-width: 0px" tabindex="1" accesskey="F"></td>
		<td nowrap>Region:</td>
		<td nowrap><select name="Region" tabindex="10">
		<%
		while (!rsRegion.EOF) {
		%>
			<option value="<%=(rsRegion.Fields.Item("insRegion_num").Value)%>" <%=((rsRegion.Fields.Item("insRegion_num").Value == rsClient.Fields.Item("insRegion_num").Value)?"SELECTED":"")%> ><%=(rsRegion.Fields.Item("chvname").Value)%></option>
		<%
			rsRegion.MoveNext();
		}
		%>
		</select></td>		
	</tr>
	<tr> 
		<td nowrap>First Name:</td>
		<td nowrap><input type="text" name="FirstName" value="<%=Trim(rsClient.Fields.Item("chvFst_name").Value)%>" tabindex="2"></td>
		<td nowrap>Status:</td>
		<td nowrap><select name="Status" style="width: 180px" tabindex="11">
		<%
		while (!rsStatus.EOF) {
		%>
			<option value="<%=(rsStatus.Fields.Item("insStdnt_status_id").Value)%>" <%=((rsStatus.Fields.Item("insStdnt_status_id").Value == rsClient.Fields.Item("insStdnt_Status_id").Value)?"SELECTED":"")%> ><%=(rsStatus.Fields.Item("chvName").Value)%></option>
		<%
			rsStatus.MoveNext();
		}
		%>
		</select></td>		
	</tr>
	<tr> 
		<td nowrap>Middle Name:</td>
		<td nowrap><input type="text" name="MiddleName" value="<%=Trim(rsClient.Fields.Item("chvMdl_name").Value)%>" maxlength="50" tabindex="3"></td>
		<td nowrap>Case Manager:</td>
		<td nowrap><select name="CaseManager" tabindex="12" style="width: 180px">
		<%
		while (!rsCaseManager.EOF) {
		%>
			<option value="<%=(rsCaseManager.Fields.Item("insId").Value)%>" <%=((rsCaseManager.Fields.Item("insId").Value == rsClient.Fields.Item("insCase_mngr_id").Value)?"SELECTED":"")%> ><%=(rsCaseManager.Fields.Item("chvName").Value)%></option>
		<%
			rsCaseManager.MoveNext();
		}
		%>
		</select></td>				
	</tr>
    <tr> 
		<td nowrap>Last Name:</td>
		<td nowrap><input type="text" name="LastName" value="<%=Trim(rsClient.Fields.Item("chvLst_name").Value)%>" maxlength="50" tabindex="4"></td>
		<td nowrap>Primary Disability:</td>
		<td nowrap><select name="PrimaryDisability" tabindex="13" style="width: 180px">
		<%
		while (!rsDisability.EOF) {
		%>
			<option value="<%=(rsDisability.Fields.Item("insDisability_id").Value)%>" <%=((rsDisability.Fields.Item("insDisability_id").Value == rsClient.Fields.Item("insDsbty1_id").Value)?"SELECTED":"")%> ><%=(rsDisability.Fields.Item("chvname").Value)%></option>
		<%
			rsDisability.MoveNext();
		}
		%>
		</select></td>				
	</tr>
	<tr> 
		<td nowrap>SIN:</td>
		<td nowrap><input type="text" name="Sin" value="<%=FormatSIN(Trim(rsClient.Fields.Item("chrSIN_no").Value))%>" size="15" maxlength="11" tabindex="5" onChange="FormatSIN(this);" ></td>
		<td nowrap>Secondary Disability:</td>
		<td nowrap><select name="SecondaryDisability" tabindex="14" style="width: 180px">
		<%
		rsDisability.MoveFirst();
		while (!rsDisability.EOF) {
		%>
			<option value="<%=(rsDisability.Fields.Item("insDisability_id").Value)%>" <%=((rsDisability.Fields.Item("insDisability_id").Value == rsClient.Fields.Item("insDsbty2_id").Value)?"SELECTED":"")%>><%=(rsDisability.Fields.Item("chvname").Value)%></option>
		<%
			rsDisability.MoveNext();
		}
		%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>PEN:</td>
		<td nowrap><input type="text" name="Pen" value="<%=Trim(rsClient.Fields.Item("chrPEN_num").Value)%>" size="15" maxlength="9" tabindex="6" onKeypress="AllowNumericOnly();"></td>
		<td nowrap>Program Standing:</td>
		<td nowrap><select name="ProgramStanding" style="width: 180px" tabindex="15">
			<option value="0" <%=((rsClient.Fields.Item("bitIsDefault_asp").Value == 0)?"SELECTED":"")%>>In Good Standing
			<option value="1" <%=((rsClient.Fields.Item("bitIsDefault_asp").Value == 1)?"SELECTED":"")%>>Default
		</select></td> 		
	</tr>
	<tr>
		<td nowrap>Gender:</td>
		<td nowrap><select name="Gender" tabindex="7">
			<option value="1" <%=((rsClient.Fields.Item("bitGender_is_male").Value == 1)?"Selected":"")%>>Male
			<option value="0" <%=((rsClient.Fields.Item("bitGender_is_male").Value == 0)?"Selected":"")%>>Female
		</select></td>		
		<td nowrap>Past Service Received:</td>
		<td nowrap>
			<input type="checkbox" name="SetBCServed" <%=((rsClient.Fields.Item("bitIs_Prx_SETBC").Value == 1)?"CHECKED":"")%> value="1" tabindex="16" class="chkstyle">SetBC
	        <input type="checkbox" name="PRCVIServed" <%=((rsClient.Fields.Item("bitIs_Prx_PRCVI").Value == 1)?"CHECKED":"")%> value="1" tabindex="17" class="chkstyle">PRCVI
		</td>
    </tr>
    <tr> 
		<td nowrap>Date of Birth:</td>
		<td nowrap>
			<input type="text" name="DateOfBirth" value="<%=FilterDate(rsClient.Fields.Item("dtsBirth_date").Value)%>" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Is First Nation:</td>
		<td nowrap><select name="IsFirstNation" tabindex="18" accesskey="L">
			<option value="0" <%=((rsClient.Fields.Item("bitIs_FirstNations").Value == 0)?"SELECTED":"")%>>No
			<option value="1" <%=((rsClient.Fields.Item("bitIs_FirstNations").Value == 1)?"SELECTED":"")%>>Yes
		</select></td>	
	</tr>
	<tr>
		<td nowrap>Age:</td>
		<td nowrap><input type="text" name="Age" value="<%=rsClient.Fields.Item("intAge").Value%>" tabindex="9" size="3" maxlength="3" readonly></td>
		<td colspan="2"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="19" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="20" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsClient.Fields.Item("intAdult_id").Value%>">
</form>
</body>
</html>
<%
rsRegion.Close();
rsStatus.Close();
rsDisability.Close();
rsCaseManager.Close();
rsClient.Close();
%>
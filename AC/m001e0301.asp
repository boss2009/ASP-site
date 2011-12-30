<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
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
  var MM_editTable  = "dbo.tbl_Address";
  var MM_editColumn = "intAddress_id";
  var MM_recordId = "" + Request.Form("MM_recordId") + "";
  var MM_editRedirectUrl = "UpdateSuccessful.asp?page=m001q0301.asp&intAdult_id="+Request.QueryString("intAdult_id");
  var MM_fieldsStr = "StreetAddress|value|City|value|Province|value|PostalCode|value|PrimaryPhoneType|value|PrimaryPhoneAreaCode|value|PrimaryPhoneNumber|value|PrimaryPhoneExtension|value|SecondaryPhoneType|value|SecondaryPhoneAreaCode|value|SecondaryPhoneNumber|value|SecondaryPhoneExtension|value|EMail|value|Notes|value";
  var MM_columnsStr = "chvAddress|',none,''|chvCity|',none,''|insProv_State_id|none,none,NULL|chvPostal_zip|',none,''|intPhone_Type_1|none,none,NULL|chvPhone1_Arcd|',none,''|chvPhone1_Num|',none,''|chvPhone1_Ext|',none,''|intPhone_Type_2|none,none,NULL|chvPhone2_Arcd|',none,''|chvPhone2_Num|',none,''|chvPhone2_Ext|',none,''|chvEmail|',none,''|chvNote|',none,''";

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

var rsAddress = Server.CreateObject("ADODB.Recordset");
rsAddress.ActiveConnection = MM_cnnASP02_STRING;
rsAddress.Source = "{call dbo.cp_Idv_Adult_Address("+ Request.QueryString("intaddr_id") + ")}";
rsAddress.CursorType = 0;
rsAddress.CursorLocation = 2;
rsAddress.LockType = 3;
rsAddress.Open();

var rsProvince  = Server.CreateObject("ADODB.Recordset");
rsProvince .ActiveConnection = MM_cnnASP02_STRING;
rsProvince .Source = "{call dbo.cp_Prov_State}";
rsProvince .CursorType = 0;
rsProvince .CursorLocation = 2;
rsProvince .LockType = 3;
rsProvince .Open();

var rsPhoneType = Server.CreateObject("ADODB.Recordset");
rsPhoneType.ActiveConnection = MM_cnnASP02_STRING;
rsPhoneType.Source = "{call dbo.cp_Phone_Type}";
rsPhoneType.CursorType = 0;
rsPhoneType.CursorLocation = 2;
rsPhoneType.LockType = 3;
rsPhoneType.Open();

var rsAreaCode = Server.CreateObject("ADODB.Recordset");
rsAreaCode.ActiveConnection = MM_cnnASP02_STRING;
rsAreaCode.Source = "{call dbo.cp_area_code(0,'',0,2,'Q',0)}";
rsAreaCode.CursorType = 0;
rsAreaCode.CursorLocation = 2;
rsAreaCode.LockType = 3;
rsAreaCode.Open();
%>									
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>Update Address</title>
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
				document.frm0301.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (!CheckTextArea(document.frm0301.Notes, 50)){
			alert("Text area cannot exceed 50 characters.");
			return ;
		}
		if (!CheckPostalCode(document.frm0301.PostalCode.value)){
			alert("Invalid Postal Code.");
			document.frm0301.PostalCode.focus();
			return ;
		}
		if (!CheckEmail(document.frm0301.EMail.value)){
			alert("Invalid Email.");
			document.frm0301.EMail.focus();
			return ;
		}
		var tempPC = document.frm0301.PostalCode.value;
		tempPC = tempPC.toUpperCase();
		document.frm0301.PostalCode.value = tempPC;		
				
		document.frm0301.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0301.StreetAddress.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0301">
<h5>Update Address</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap valign="top">Street Address:</td>
		<td nowrap valign="top"><textarea name="StreetAddress" cols="30" rows="3" tabindex="1" accesskey="F"><%=Trim(rsAddress.Fields.Item("chvAddress").Value)%></textarea></td>
	</tr>
	<tr> 
		<td nowrap>City:</td>
		<td nowrap><input type="text" name="City" value="<%=Trim(rsAddress.Fields.Item("chvCity").Value)%>" maxlength="50" tabindex="2"></td>
	</tr>
	<tr> 
		<td nowrap>Province:</td>
		<td nowrap><select name="Province" tabindex="3">
		<%
		while (!rsProvince .EOF) {
		%>
			<option value="<%=(rsProvince .Fields.Item("intprvst_id").Value)%>" <%=((rsProvince .Fields.Item("intprvst_id").Value == rsAddress.Fields.Item("insProv_State_id").Value)?"SELECTED":"")%>><%=(rsProvince .Fields.Item("chrprvst_abbv").Value)%></option>
		<%
			rsProvince .MoveNext();
		}
		%>
        </select></td>
    </tr>	
    <tr> 
		<td nowrap>Postal Code:</td>
		<td nowrap><input type="text" name="PostalCode" value="<%=FormatPostalCode(Trim(rsAddress.Fields.Item("chvPostal_zip").Value))%>" tabindex="4" size="10" maxlength="7" onChange="FormatPostalCode(this);"></td>
    </tr>
    <tr> 
		<td nowrap>Primary Phone:</td>
		<td nowrap> 
			<select name="PrimaryPhoneType" tabindex="5">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%=((rsPhoneType.Fields.Item("intPhone_type_id").Value == rsAddress.Fields.Item("intPhone_Type_1").Value)?"SELECTED":"")%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			rsPhoneType.MoveFirst();
			%>
			</select>
			<select name="PrimaryPhoneAreaCode" tabindex="6">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%=((rsAddress.Fields.Item("chvPhone1_Arcd").Value == rsAreaCode.Fields.Item("chvAC_num").Value)?"SELECTED":"")%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			rsAreaCode.MoveFirst();
			%>			
			</select>
			<input type="text" name="PrimaryPhoneNumber" value="<%=FormatPhoneNumberOnly(rsAddress.Fields.Item("chvPhone1_Num").Value)%>" size="9" tabindex="7" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">Ext
			<input type="text" name="PrimaryPhoneExtension" value="<%=Trim(rsAddress.Fields.Item("chvPhone1_Ext").Value)%>" size="4" tabindex="8" onKeypress="AllowNumericOnly();">
		</td>
    </tr>
    <tr> 
		<td nowrap>Secondary Phone:</td>
		<td nowrap>
			<select name="SecondaryPhoneType" tabindex="9">
			<% 
			while (!rsPhoneType.EOF) {
			%>
				<option value="<%=(rsPhoneType.Fields.Item("intPhone_type_id").Value)%>" <%=((rsPhoneType.Fields.Item("intPhone_type_id").Value == rsAddress.Fields.Item("intPhone_Type_2").Value)?"SELECTED":"")%>><%=(rsPhoneType.Fields.Item("chvName").Value)%></option>
			<%
				rsPhoneType.MoveNext();
			}
			%>
			</select>
			<select name="SecondaryPhoneAreaCode" tabindex="10">
			<%
			while (!rsAreaCode.EOF) {			
			%>
				<option value="<%=rsAreaCode.Fields.Item("chvAC_num").Value%>" <%=((rsAddress.Fields.Item("chvPhone2_Arcd").Value == rsAreaCode.Fields.Item("chvAC_num").Value)?"SELECTED":"")%>><%=rsAreaCode.Fields.Item("chvAC_num").Value%>
			<%
				rsAreaCode.MoveNext();
			}
			%>			
			</select>
			<input type="text" name="SecondaryPhoneNumber" value="<%=FormatPhoneNumberOnly(rsAddress.Fields.Item("chvPhone2_Num").Value)%>" size="9" tabindex="11" onKeypress="AllowNumericOnly();" maxlength="8" onChange="FormatPhoneNumberOnly(this)">Ext
			<input type="text" name="SecondaryPhoneExtension" value="<%=Trim(rsAddress.Fields.Item("chvPhone2_Ext").Value)%>" size="4" tabindex="12" onKeypress="AllowNumericOnly();" >
		</td>
    </tr>
    <tr> 
		<td nowrap>E-Mail:</td>
		<td nowrap><input type="text" name="EMail" value="<%=Trim(rsAddress.Fields.Item("chvEmail").Value)%>" tabindex="13"></td>
	</tr>  
	<tr>
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="14" accesskey="L"><%=Trim(rsAddress.Fields.Item("chvNote").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="16" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="17" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=(rsAddress.Fields.Item("intAddress_id").Value)%>">
</form>
</body>
</html>
<%
rsAddress.Close();
rsProvince .Close();
rsPhoneType.Close();
rsAreaCode.Close();
%>
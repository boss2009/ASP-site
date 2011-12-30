<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
if(String(Request.Form("MM_flag")) == "1"){

   var chvTemp = "MM_ID: "+Request.Form("MM_ID") + "<P>" +
                 "MM_buffer: "+Request.Form("MM_buffer") + "<P>" + 
                 "Description: "+Request.Form("Description") + "<P>" + 
                 "IsActive: "+Request.Form("IsActive") + "<P>" + 
                 "IsLoan:   "+Request.Form("IsLoan") + "<P>" + 
                 "IsBuyOut: "+Request.Form("IsBuyOut") + "<P>" 
//   Response.Write(chvTemp);
//concatenate parameters 
   var MM_editRedirectUrl  = "m018a0312B.asp?insrefer_agent_id=" + Request.Form("MM_ID") ;
       MM_editRedirectUrl += "&chvname=" + Request.Form("Description");
// Is active
       if (String(Request.Form("IsActive")) == "1") {
          MM_editRedirectUrl += "&bitis_active=" + Request.Form("IsActive");
	   } else {
          MM_editRedirectUrl += "&bitis_active=" + "0";
	   }
// Is Loan
       if (String(Request.Form("IsLoan")) == "1") {
          MM_editRedirectUrl += "&bitis_loan="   + Request.Form("IsLoan");
	   } else {
          MM_editRedirectUrl += "&bitis_loan="   + "0";
	   }
// Is Buyout
       if (String(Request.Form("IsBuyOut")) == "1") {
          MM_editRedirectUrl += "&bitis_BuyOut=" + Request.Form("IsBuyOut");
	   } else {
          MM_editRedirectUrl += "&bitis_BuyOut=" + "0";
	   }

       MM_editRedirectUrl += "&chrFS_chbx="   + Request.Form("MM_buffer");

     if(String(Request.Form("MM_buffer")) == "" ){
// case user did not click any checkbox before submit
		Response.Write("<P>");	 
        Response.Write("No Add is allowed if no check box is clicked ...");
		Response.Write("<P>");	 
	 } else {
//	    Response.Write("MM_editRedirectUrl : "+MM_editRedirectUrl); 
        Response.Redirect(MM_editRedirectUrl);
	 }

}

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_funding_source(0,0)}";
rsFundingSource.CursorType = 0;          
rsFundingSource.CursorLocation = 2;      
rsFundingSource.LockType = 3;              
rsFundingSource.Open();
var rsFundingSource_numRows = 0;
%>
<%
      var chrFSOption = "00000000000000000000000000000000000000000000000000";
//      Response.Write("Option is: "+chrFSOption);
   
      var FSArray = new Array();
      var FSMaxLen = chrFSOption.length ;
      for (var i=0;i < FSMaxLen; i++){
         FSArray[i] = chrFSOption.substr(i,1);
//	     Response.Write("index: "+i+" content: "+FSArray[i]+"<P>");
      }

%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsFundingSource_numRows += Repeat1__numRows;

%>
<html>
<head>
	<title>New Referral Type</title>
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
		if (Trim(document.frm0312.Description.value)=="") {
			alert("Enter Referral Type Description.");
			document.frm0312.Description.focus();
			return ;
		}
		var len = document.frm0312.elements.length;
		var buffer = "";
		var y = 0;
		var wdest = ""
		
		for (var i=0;i<len; i++){
			if (document.frm0312.elements[i].name == "FundingSourceCheckBox") {
				if (document.frm0312.elements[i].checked) {
					buffer += "1" ; 
				} else {
					buffer += "0" ; 
				}
				y += 1;
			}
		}
		document.frm0312.MM_buffer.value = buffer;
		document.frm0312.action = "m018a0312.asp";
		document.frm0312.submit();
	}
</Script>
</head>
<body onLoad="document.frm0312.Description.focus();">
<form name="frm0312" method="post" action="">
<h5>New Referral Type</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="30" accesskey="F" ></td>
    </tr>
    <tr>
		<td>Is Active:</td> 
		<td><input type="checkbox" name="IsActive" value="1" class="chkstyle"></td>
	</tr>
	<tr>
		<td>Is Loan:</td>
		<td><input type="checkbox" name="IsLoan" value="1" class="chkstyle"></td>
	</tr>
	<tr>
		<td>Is Buyout:</td> 
		<td><input type="checkbox" name="IsBuyOut" value="1" accesskey="L" class="chkstyle"></td>
    </tr>
    <tr> 
      	<td colspan="2">Funding Source:</td>
    </tr>
    <% 
	while ((Repeat1__numRows-- != 0) && (!rsFundingSource.EOF)) { %>
	<tr>
		<td><input type="checkbox" name="FundingSourceCheckBox" <%=((FSArray[Repeat1__index] == 1)?"CHECKED":"")%> value="<%=(rsFundingSource.Fields.Item("insFunding_source_id").Value)%>" class="chkstyle"><%=(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%></td>
	<% 
		Repeat1__index++;
		rsFundingSource.MoveNext();
		if(!rsFundingSource.EOF) {
	%>
		<td><input type="checkbox" name="FundingSourceCheckBox" <%=((FSArray[Repeat1__index] == 1)?"CHECKED":"")%> value="<%=(rsFundingSource.Fields.Item("insFunding_source_id").Value)%>" class="chkstyle"><%=(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%></td>
	<%
		} else {
			Response.Write("&nbsp;");
		}
	%>
	</tr>
	<%
		Repeat1__index++;
		rsFundingSource.MoveNext();
	}
	%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_buffer">
<input type="hidden" name="MM_flag" value="1">
<input type="hidden" name="MM_ID" value="">
</form>
<% Response.Write("Total Num of Funding Source: "+Repeat1__index); %>
</body>
</html>
<%
rsFundingSource.Close();
%>

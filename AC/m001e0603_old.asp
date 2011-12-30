<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update")) == "true"){
	var rsGrantEligibility = Server.CreateObject("ADODB.Recordset");
	rsGrantEligibility.ActiveConnection = MM_cnnASP02_STRING;
	rsGrantEligibility.Source = "{call dbo.cp_grant_elgbty3("+ Request.Form("ReferralDate") + ","+Request.QueryString("intAdult_id")+","+Request.Form("GrantPeriod")+",0,'E',0)}";
	rsGrantEligibility.CursorType = 0;
	rsGrantEligibility.CursorLocation = 2;
	rsGrantEligibility.LockType = 3;
	rsGrantEligibility.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m001q0603.asp&intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsGrantEligibility = Server.CreateObject("ADODB.Recordset");
rsGrantEligibility.ActiveConnection = MM_cnnASP02_STRING;
rsGrantEligibility.Source = "{call dbo.cp_grant_elgbty3("+ Request.QueryString("intReferral_id") + ","+Request.QueryString("intAdult_id")+",0,1,'Q',0)}";
rsGrantEligibility.CursorType = 0;
rsGrantEligibility.CursorLocation = 2;
rsGrantEligibility.LockType = 3;
rsGrantEligibility.Open();

var rsGrantPeriod = Server.CreateObject("ADODB.Recordset");
rsGrantPeriod.ActiveConnection = MM_cnnASP02_STRING;
rsGrantPeriod.Source = "{call dbo.cp_list_elgbty_period2(0,"+ Request.QueryString("intAdult_id") + ",0,'','',0.0,0.0,0.0,0,'',0,'',0,'Q',0)}";
rsGrantPeriod.CursorType = 0;
rsGrantPeriod.CursorLocation = 2;
rsGrantPeriod.LockType = 3;
rsGrantPeriod.Open();

var rsReferralDate = Server.CreateObject("ADODB.Recordset");
rsReferralDate.ActiveConnection = MM_cnnASP02_STRING;
rsReferralDate.Source = "{call dbo.cp_list_referrals("+ Request.QueryString("intAdult_id") + ",0,0,1,0)}";
rsReferralDate.CursorType = 0;
rsReferralDate.CursorLocation = 2;
rsReferralDate.LockType = 3;
rsReferralDate.Open();
var rsReferralDate_Total = 0
while (!rsReferralDate.EOF) {
	rsReferralDate_Total++;
	rsReferralDate.MoveNext();
}
if (rsReferralDate_Total > 0) rsReferralDate.MoveFirst();

var ChkGrantEligibility = Server.CreateObject("ADODB.Command");
ChkGrantEligibility.ActiveConnection = MM_cnnASP02_STRING;
ChkGrantEligibility.CommandText = "dbo.cp_Chk_Grant_Elgbty";
ChkGrantEligibility.CommandType = 4;
ChkGrantEligibility.CommandTimeout = 0;
ChkGrantEligibility.Prepared = true;
ChkGrantEligibility.Parameters.Append(ChkGrantEligibility.CreateParameter("RETURN_VALUE", 3, 4));
ChkGrantEligibility.Parameters.Append(ChkGrantEligibility.CreateParameter("@intAdult_id", 3, 1,10000,Request.QueryString("intAdult_id")));
ChkGrantEligibility.Parameters.Append(ChkGrantEligibility.CreateParameter("@insRtnFlag", 2, 2));
ChkGrantEligibility.Execute();
%>
<html>
<head>
	<title>Update Grant Eligibility</title>
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
				document.frm0603.reset();
			break;
		   	case 76 :
				//alert("L");
				window.location.href='m001q0603.asp?<%=Request.QueryString%>';
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, name, "width=600,height=570,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	
	var ReferralArray = new Array(<%=rsReferralDate_Total%>);
		ReferralArray[0] = new Array(6);
		ReferralArray[0][0] = 0;
		ReferralArray[0][1] = 0;
		ReferralArray[0][2] = 0;
		ReferralArray[0][3] = 0;
		ReferralArray[0][4] = 0;		
		ReferralArray[0][5] = 0;	
	<%
	var i = 1;
	while (!rsReferralDate.EOF) {
	%>
		ReferralArray[<%=i%>] = new Array(6);
		ReferralArray[<%=i%>][0] = <%=rsReferralDate.Fields.Item("intReferral_id").Value%>;
		ReferralArray[<%=i%>][1] = <%=((rsReferralDate.Fields.Item("bitIs_Re-referral")=="1")?"1":"0")%>;
		ReferralArray[<%=i%>][2] = <%=((rsReferralDate.Fields.Item("bitIs_PS_CSG_Grant")=="1")?"1":"0")%>;
		ReferralArray[<%=i%>][3] = <%=((rsReferralDate.Fields.Item("bitIs_PS_APSD_Grant")=="1")?"1":"0")%>;
		ReferralArray[<%=i%>][4] = <%=((rsReferralDate.Fields.Item("bitIs_VRS_CSG_Grant")=="1")?"1":"0")%>;		
		ReferralArray[<%=i%>][5] = <%=((rsReferralDate.Fields.Item("insGrnt_Elgbty").Value>0)?rsReferralDate.Fields.Item("insGrnt_Elgbty").Value:"0")%>;
	<%
		i++;
		rsReferralDate.MoveNext();
	}
	if (rsReferralDate_Total>0) rsReferralDate.MoveFirst();
	%>
	function Save(){
		if (document.frm0603.ReferralDate.value==0) {
			alert("Select a referral date before linking to grant period.");
			return ;
		}
		var j = 0
		if (String(document.frm0603.LinkToReferral.length)=="undefined") {
			if (document.frm0603.LinkToReferral.checked) j++;
		} else {
			for (var i=0; i< document.frm0603.LinkToReferral.length; i++) {
				if (document.frm0603.LinkToReferral[i].checked) j++;
			}
		}
		if (j==0) {
			alert("Select a grant study period.");
			return ;
		}
		
		if (j>1) {
			alert("A referral can only have one grant period assigned.");
			return ;
		}
		document.frm0603.GrantPeriod.value=0;
		if (String(document.frm0603.LinkToReferral.length)=="undefined") {
			if (document.frm0603.LinkToReferral.checked) document.frm0603.GrantPeriod.value=document.frm0603.LinkToReferral.value;
		} else {
			for (var i=0; i< document.frm0603.LinkToReferral.length; i++) {
				if (document.frm0603.LinkToReferral[i].checked) document.frm0603.GrantPeriod.value=document.frm0603.LinkToReferral[i].value;
			}
		}	
		document.frm0603.submit();
	}
	
	function Init(){
		document.frm0603.ReferralDate.focus();
		ChangeReferralDate();
	}
	
	function SetGrantPeriod(grant_id){
	<%
	if (!rsGrantPeriod.EOF) {
	%>
		if (String(document.frm0603.LinkToReferral.length)=="undefined") {
			if (document.frm0603.LinkToReferral.value==grant_id) {
				document.frm0603.LinkToReferral.checked=true;
			} else {
				document.frm0603.LinkToReferral.checked=false;
			}			
		} else {
			for (var i=0; i< document.frm0603.LinkToReferral.length; i++) {
				if (document.frm0603.LinkToReferral[i].value==grant_id) {
					document.frm0603.LinkToReferral[i].checked=true;
				} else {
					document.frm0603.LinkToReferral[i].checked=false;
				}
			}
		}
	<%
	}
	%>
	}
	
	function ChangeReferralDate(){
		document.frm0603.ReferralType.value=ReferralArray[document.frm0603.ReferralDate.selectedIndex][1];
		document.frm0603.PostSecondaryCSG.checked=ReferralArray[document.frm0603.ReferralDate.selectedIndex][2];
		document.frm0603.PostSecondaryAPSD.checked=ReferralArray[document.frm0603.ReferralDate.selectedIndex][3];
		document.frm0603.EPPDCSG.checked=ReferralArray[document.frm0603.ReferralDate.selectedIndex][4];
		SetGrantPeriod(ReferralArray[document.frm0603.ReferralDate.selectedIndex][5]);
	}
	</script>
</head>
<body onLoad="Init();">
<form name="frm0603" method="POST" action="<%=MM_editAction%>">
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral Date:</td>
		<td nowrap><select name="ReferralDate" acccesskey="F" tabindex="1" onChange="ChangeReferralDate();">
				<option value="0">(not linked)
			<%
			while (!rsReferralDate.EOF){
			%>
				<option value="<%=rsReferralDate.Fields.Item("intReferral_id").Value%>" <%=((rsReferralDate.Fields.Item("intReferral_id").Value==rsGrantEligibility.Fields.Item("intReferral_id").Value)?"SELECTED":"")%>><%=rsReferralDate.Fields.Item("dtsRefral_date").Value%>
			<%
				rsReferralDate.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>Referral Type:</td>
		<td nowrap><select name="ReferralType" tabindex="2">
			<option value="0">Referral
			<option value="1">Re-referral
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Grant Type:</td>
		<td nowrap>
	  		<input type="checkbox" name="PostSecondaryCSG" value="1" tabindex="3" class="chkstyle">PS-CSG
			<input type="checkbox" name="PostSecondaryAPSD" value="1" tabindex="4" class="chkstyle">PS-APSD
			<input type="checkbox" name="EPPDCSG" value="1" tabindex="5" class="chkstyle">EPPD-CSG
		</td>
    </tr>
</table><br>
<b>Grant Study Periods:</b>
<table cellpadding="1" cellspacing="1" class="MTable">
	<tr>
		<th class="headrow" nowrap>Link To Referral</th>
		<th class="headrow" nowrap>Funding Source</th>		
		<th class="headrow" nowrap>Start Date</th>
		<th class="headrow" nowrap>End Date</th>
		<th class="headrow" nowrap>Grant Amount</th>
	</tr>
<%
var count = 0;
while (!rsGrantPeriod.EOF) {
%>	
	<tr>
		<td align="center"><input type="checkbox" name="LinkToReferral" value="<%=rsGrantPeriod.Fields.Item("insGrnt_Elgbty").Value%>" class="chkstyle" style="background-color: #ffffe6"></td>
<!-- // + Nov.03.2005
		<td align="center"><a href="m001e0603b.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intReferral_id=<%=rsGrantEligibility.Fields.Item("intReferral_id").Value%>&insGrnt_Elgbty=<%=rsGrantPeriod.Fields.Item("insGrnt_Elgbty").Value%>"><%=(rsGrantPeriod.Fields.Item("chvfunding_source_name").Value)%></a></td>		
-->
		<td align="center"><a href="m001e0603b.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intReferral_id=<%=rsGrantEligibility.Fields.Item("intReferral_id").Value%>&insGrnt_Elgbty=<%=rsGrantPeriod.Fields.Item("insGrnt_Elgbty").Value%>"><%=(rsGrantPeriod.Fields.Item("chvfunding_source_name").Value)%></a></td>		

		<td align="center"><%=FilterDate(rsGrantPeriod.Fields.Item("dtmEligibility_start").Value)%></td>
		<td align="center"><%=FilterDate(rsGrantPeriod.Fields.Item("dtmEligibility_end").Value)%></td>
		<td align="right"><%=FormatCurrency(rsGrantPeriod.Fields.Item("fltGrantAmt").Value)%></td>
	</tr>
<%
	count++;
	rsGrantPeriod.MoveNext();
}
%>
</table>
<br><br>
<a href="javascript: openWindow('m001e0603a.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>','W0603E');">Add Financial Eligibility Information</a>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" <%=((count==0)?"DISABLED":"")%> class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m001q0603.asp?<%=Request.QueryString%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="GrantPeriod" value="">
</form>
</body>
</html>
<%
rsGrantPeriod.Close();
rsGrantEligibility.Close();
rsReferralDate.Close();
%>
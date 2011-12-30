<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.Form("MM_action"))=="Remove") {
	var rsRemoveComponent = Server.CreateObject("ADODB.Recordset");
	rsRemoveComponent.ActiveConnection = MM_cnnASP02_STRING;
	rsRemoveComponent.Source = "{call dbo.cp_bundle_eqp_class("+Request.QueryString("insBundle_id")+","+Request.QueryString("insEquip_Class_id")+",0,'D',0)}";
	rsRemoveComponent.CursorType = 0;
	rsRemoveComponent.CursorLocation = 2;
	rsRemoveComponent.LockType = 3;
	rsRemoveComponent.Open();
	Response.Redirect("UpdateSuccessful.asp");
}

if (String(Request.Form("MM_action"))=="Add") {
	var rsAddComponent = Server.CreateObject("ADODB.Recordset");
	rsAddComponent.ActiveConnection = MM_cnnASP02_STRING;
	rsAddComponent.Source = "{call dbo.cp_bundle_eqp_class("+Request.QueryString("insBundle_id")+","+Request.Form("ClassID")+",1,'A',0)}";
	rsAddComponent.CursorType = 0;
	rsAddComponent.CursorLocation = 2;
	rsAddComponent.LockType = 3;
	rsAddComponent.Open();
	Response.Redirect("UpdateSuccessful.asp");	
}

var rsComponent = Server.CreateObject("ADODB.Recordset");
rsComponent.ActiveConnection = MM_cnnASP02_STRING;
rsComponent.Source = "{call dbo.cp_bundle_eqp_class("+Request.QueryString("insBundle_id")+",0,0,'Q',0)}";
rsComponent.CursorType = 0;
rsComponent.CursorLocation = 2;
rsComponent.LockType = 3;
rsComponent.Open();
%>
<html>
<head>
	<title>Bundle Components</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	
	function AddComponent(){
		if (document.frm05s01.ClassID.value=="") return ;
		document.frm05s01.MM_action.value="Add";
		document.frm05s01.action = "m005e0102.asp?insBundle_id=<%=Request.QueryString("insBundle_id")%>";
		document.frm05s01.submit();	
	}
	
	function RemoveComponent(bundle_id, class_id){
		document.frm05s01.MM_action.value="Remove";
		document.frm05s01.action = "m005e0102.asp?insBundle_id="+bundle_id+"&insEquip_Class_id="+class_id;
		document.frm05s01.submit();
	}
	</Script>
</head>
<body>
<form name="frm05s01" method="post" action="">
<b>Components</b><br>
  <div class="BrowsePanel" style="height: 240px; width: 460px"> 
    <table cellpadding="2" cellspacing="1" class="Mtable" width="440">
      <tr> 
        <th nowrap class="headrow" align="left">Inventory Class</th>
        <th nowrap class="headrow" align="left">List Unit Cost</th>
        <th nowrap class="headrow" align="center">&nbsp;</th>
      </tr>
      <% 
var Total = 0;
while (!rsComponent.EOF) { 
%>
      <tr> 
        <td nowrap valign="top"><a href="javascript: openWindow('../EC/m007FS3.asp?ClassID=<%=(rsComponent.Fields.Item("insEquip_Class_id").Value)%>','wE03');"><%=(rsComponent.Fields.Item("chvEqCls_name").Value)%></a></td>
        <td nowrap valign="top" align="right" width="90"><%=FormatCurrency(rsComponent.Fields.Item("FltList_Unit_Cost").Value)%>&nbsp;</td>
        <td nowrap valign="top" align="center"><a href="javascript: RemoveComponent(<%=Request.QueryString("insBundle_id")%>, <%=(rsComponent.Fields.Item("insEquip_Class_id").Value)%>);"><img src="../i/remove.gif" ALT="Remove <%=(rsComponent.Fields.Item("chvEqCls_name").Value)%>"></a></td>
      </tr>
      <%
	Total += rsComponent.Fields.Item("FltList_Unit_Cost").Value;
	rsComponent.MoveNext();
}
%>
    </table>
  </div>
<div style="position: absolute; top: 275px; left:250px"><b>Component Total: <%=FormatCurrency(Total)%></b></div>
<div style="position: absolute; top: 300px">
Add Component:&nbsp;&nbsp;<input type="text" name="ClassName" value="" size="35" readonly>
<input type="button" value="Find" onClick="openWindow('m005p01FS.asp','FindClass')" class="btnstyle">
<input type="button" value="Add" onClick="AddComponent();" class="btnstyle">
<input type="button" value="Clear" onClick="document.frm05s01.ClassName.value='';document.frm05s01.ClassID.value='';" class="btnstyle">
<input type="hidden" name="ClassID" value="">
</div>
<input type="hidden" name="MM_action" value="">
</form>
</body>
</html>
<%
rsComponent.Close();
%>
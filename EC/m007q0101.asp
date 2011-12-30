<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsEquipmentClass = Server.CreateObject("ADODB.Recordset");
rsEquipmentClass.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentClass.Source = "{call dbo.cp_Eqp_Class(0,'C',0)}";
rsEquipmentClass.CursorType = 0;
rsEquipmentClass.CursorLocation = 2;
rsEquipmentClass.LockType = 3;
rsEquipmentClass.Open();
var rsEquipmentClass_numRows = 0;
var Repeat1__numRows = -1;
var Repeat1__index = 0;
rsEquipmentClass_numRows += Repeat1__numRows;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var rsEquipmentClass_total = rsEquipmentClass.RecordCount;

// set the number of rows displayed on this page
if (rsEquipmentClass_numRows < 0) {            // if repeat region set to all records
  rsEquipmentClass_numRows = rsEquipmentClass_total;
} else if (rsEquipmentClass_numRows == 0) {    // if no repeat regions
  rsEquipmentClass_numRows = 1;
}

// set the first and last displayed record
var rsEquipmentClass_first = 1;
var rsEquipmentClass_last  = rsEquipmentClass_first + rsEquipmentClass_numRows - 1;

// if we have the correct record count, check the other stats
if (rsEquipmentClass_total != -1) {
  rsEquipmentClass_numRows = Math.min(rsEquipmentClass_numRows, rsEquipmentClass_total);
  rsEquipmentClass_first   = Math.min(rsEquipmentClass_first, rsEquipmentClass_total);
  rsEquipmentClass_last    = Math.min(rsEquipmentClass_last, rsEquipmentClass_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (rsEquipmentClass_total == -1) {

  // count the total records by iterating through the recordset
  for (rsEquipmentClass_total=0; !rsEquipmentClass.EOF; rsEquipmentClass.MoveNext()) {
    rsEquipmentClass_total++;
  }

  // reset the cursor to the beginning
  if (rsEquipmentClass.CursorType > 0) {
    if (!rsEquipmentClass.BOF) rsEquipmentClass.MoveFirst();
  } else {
    rsEquipmentClass.Requery();
  }

  // set the number of rows displayed on this page
  if (rsEquipmentClass_numRows < 0 || rsEquipmentClass_numRows > rsEquipmentClass_total) {
    rsEquipmentClass_numRows = rsEquipmentClass_total;
  }

  // set the first and last displayed record
  rsEquipmentClass_last  = Math.min(rsEquipmentClass_first + rsEquipmentClass_numRows - 1, rsEquipmentClass_total);
  rsEquipmentClass_first = Math.min(rsEquipmentClass_first, rsEquipmentClass_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = rsEquipmentClass;
var MM_rsCount   = rsEquipmentClass_total;
var MM_size      = rsEquipmentClass_numRows;
var MM_uniqueCol = "";
    MM_paramName = "";
var MM_offset = 0;
var MM_atTotal = false;
var MM_paramIsDefined = (MM_paramName != "" && String(Request(MM_paramName)) != "undefined");
%>
<%
// *** Move To Record: handle 'index' or 'offset' parameter

if (!MM_paramIsDefined && MM_rsCount != 0) {

  // use index parameter if defined, otherwise use offset parameter
  r = String(Request("index"));
  if (r == "undefined") r = String(Request("offset"));
  if (r && r != "undefined") MM_offset = parseInt(r);

  // if we have a record count, check if we are past the end of the recordset
  if (MM_rsCount != -1) {
    if (MM_offset >= MM_rsCount || MM_offset == -1) {  // past end or move last
      if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount % MM_size);
      } else {
        MM_offset = MM_rsCount - MM_size;
      }
    }
  }

  // move the cursor to the selected record
  for (var i=0; !MM_rs.EOF && (i < MM_offset || MM_offset == -1); i++) {
    MM_rs.MoveNext();
  }
  if (MM_rs.EOF) MM_offset = i;  // set MM_offset to the last possible record
}
%>
<%
// *** Move To Record: if we dont know the record count, check the display range

if (MM_rsCount == -1) {

  // walk to the end of the display range for this page
  for (var i=MM_offset; !MM_rs.EOF && (MM_size < 0 || i < MM_offset + MM_size); i++) {
    MM_rs.MoveNext();
  }

  // if we walked off the end of the recordset, set MM_rsCount and MM_size
  if (MM_rs.EOF) {
    MM_rsCount = i;
    if (MM_size < 0 || MM_size > MM_rsCount) MM_size = MM_rsCount;
  }

  // if we walked off the end, set the offset based on page size
  if (MM_rs.EOF && !MM_paramIsDefined) {
    if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
      MM_offset = MM_rsCount - (MM_rsCount % MM_size);
    } else {
      MM_offset = MM_rsCount - MM_size;
    }
  }

  // reset the cursor to the beginning
  if (MM_rs.CursorType > 0) {
    if (!MM_rs.BOF) MM_rs.MoveFirst();
  } else {
    MM_rs.Requery();
  }

  // move the cursor to the selected record
  for (var i=0; !MM_rs.EOF && i < MM_offset; i++) {
    MM_rs.MoveNext();
  }
}
%>
<%
// *** Move To Record: update recordset stats

// set the first and last displayed record
rsEquipmentClass_first = MM_offset + 1;
rsEquipmentClass_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  rsEquipmentClass_first = Math.min(rsEquipmentClass_first, MM_rsCount);
  rsEquipmentClass_last  = Math.min(rsEquipmentClass_last, MM_rsCount);
}

// set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount != -1 && MM_offset + MM_size >= MM_rsCount);
%>
<%
// *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

// create the list of parameters which should not be maintained
var MM_removeList = "&index=";
if (MM_paramName != "") MM_removeList += "&" + MM_paramName.toLowerCase() + "=";
var MM_keepURL="",MM_keepForm="",MM_keepBoth="",MM_keepNone="";

// add the URL parameters to the MM_keepURL string
for (var items=new Enumerator(Request.QueryString); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepURL += "&" + items.item() + "=" + Server.URLencode(Request.QueryString(items.item()));
  }
}

// add the Form variables to the MM_keepForm string
for (var items=new Enumerator(Request.Form); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepForm += "&" + items.item() + "=" + Server.URLencode(Request.Form(items.item()));
  }
}

// create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL + MM_keepForm;
if (MM_keepBoth.length > 0) MM_keepBoth = MM_keepBoth.substring(1);
if (MM_keepURL.length > 0)  MM_keepURL = MM_keepURL.substring(1);
if (MM_keepForm.length > 0) MM_keepForm = MM_keepForm.substring(1);
%>
<%
// *** Move To Record: set the strings for the first, last, next, and previous links

var MM_moveFirst="",MM_moveLast="",MM_moveNext="",MM_movePrev="";
var MM_keepMove = MM_keepBoth;  // keep both Form and URL parameters for moves
var MM_moveParam = "index";

// if the page has a repeated region, remove 'offset' from the maintained parameters
if (MM_size > 1) {
  MM_moveParam = "offset";
  if (MM_keepMove.length > 0) {
    params = MM_keepMove.split("&");
    MM_keepMove = "";
    for (var i=0; i < params.length; i++) {
      var nextItem = params[i].substring(0,params[i].indexOf("="));
      if (nextItem.toLowerCase() != MM_moveParam) {
        MM_keepMove += "&" + params[i];
      }
    }
    if (MM_keepMove.length > 0) MM_keepMove = MM_keepMove.substring(1);
  }
}

// set the strings for the move to links
if (MM_keepMove.length > 0) MM_keepMove += "&";
var urlStr = Request.ServerVariables("URL") + "?" + MM_keepMove + MM_moveParam + "=";
MM_moveFirst = urlStr + "0";
MM_moveLast  = urlStr + "-1";
MM_moveNext  = urlStr + (MM_offset + MM_size);
MM_movePrev  = urlStr + Math.max(MM_offset - MM_size,0);
%>
<html>
<head>
	<title></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	function ShowClass(oItem){
		if (oItem.style.display == "block"){		
			oItem.style.display = "none";
		} else {
			oItem.style.display = "block";
		}
	}

	function EditClass(id, type){
		switch (type){
			case 'A':
				win1=window.open('m007e0101.asp?ClassID='+id,'EditAbstractClass', "width=450,height=300,scrollbars=1,left=0,top=0,status=1");
			break;
			case 'S':
				win1=window.open('m007e0102.asp?ClassID='+id,'EditSubAbstractClass', "width=450,height=300,scrollbars=1,left=0,top=0,status=1");
			break;
			case 'C':
				win1=window.open('m007FS3.asp?ClassID='+id,'EditConcreteClass', "width=700,height=600,scrollbars=1,left=0,top=0,status=1");
			break;		
		}
	}

	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=450,height=300,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	
	
	function AddClass(type){
		switch (type){
			case 'A':
				openWindow('m007a0101.asp','NewAbstractClass');
			break;
			case 'S':
				openWindow('m007a0102.asp','NewSubAbstractClass');
			break;
			case 'C':
				openWindow('m007a0103.asp','NewConcreteClass');
			break;		
		}
	}
	
	function CloseLoading(){
		parent.EquipmentClassMenuFrame.closeWindow();
		window.self.focus();
	}	
	</Script>
</head>
<body onLoad="CloseLoading();">
<h3>Equipment Classes</h3>
<a class="blue" href="javascript: AddClass('A');">Add Abstract Class</a> | <a class="green" href="javascript: AddClass('S');">Add SubAbstract Class</a> | <a class="red" href="javascript: AddClass('C');">Add Concrete Class</a><br>
<hr>
<%
var firstrow = 1;
var firstrow2 = 1;
var LastABSClass = "";
var LastSubABSClass = "";
var LastConcClass = "";
var CurrentABSClass = "";
var CurrentSubABSClass = "";
var CurrentConcClass = "";
CurrentABSClass=rsEquipmentClass.Fields.Item("insAbsCls_id").Value;
while (!rsEquipmentClass.EOF){
	CurrentABSClass=rsEquipmentClass.Fields.Item("insAbsCls_id").Value;
	if (CurrentABSClass != LastABSClass) {
		if (firstrow!=1) {
			Response.Write("</DIV><!--SubABSClass End-->\n</DIV><!--ABSClass End-->\n");
		}
%>	
		<a href="javascript: ShowClass(A<%=rsEquipmentClass.Fields.Item("insAbsCls_id").Value%>)"><img src="../i/collapse.gif" id="I<%=rsEquipmentClass.Fields.Item("insAbsCls_id").Value%>" align="absmiddle" ALT="Collapse / Expand Abstract Class <%=rsEquipmentClass.Fields.Item("chvABSClsName").Value%>"></a><a class="blue" href="javascript: EditClass('<%=rsEquipmentClass.Fields.Item("insAbsCls_id").Value%>','A');"><%=rsEquipmentClass.Fields.Item("chvABSClsName").Value%></a><br><DIV ID="A<%=rsEquipmentClass.Fields.Item("insAbsCls_id").Value%>" style="display: none">
<%
		LastSubABSClass = "";
		firstrow = 0;
		firstrow2 = 1;
	} else {
		CurrentSubABSClass=rsEquipmentClass.Fields.Item("insSubAbsCls_id").Value;
		var temp = rsEquipmentClass.Fields.Item("chvSubAbsClsName").Value;
		var temp2 = temp.split(">");
		if (CurrentSubABSClass != LastSubABSClass) {
			if (firstrow2!=1) {
				Response.Write("</DIV><!--SubABSClass End-->\n");
			}
%>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript: ShowClass(S<%=rsEquipmentClass.Fields.Item("insSubAbsCls_id").Value%>)">
			<img src="../i/collapse.gif" id="I<%=rsEquipmentClass.Fields.Item("insSubAbsCls_id").Value%>" align="absmiddle" ALT="Collapse / Expand Sub Abstract Class <%=rsEquipmentClass.Fields.Item("chvSubAbsClsName").Value%>"></a>
			<a class="green" href="javascript: EditClass('<%=rsEquipmentClass.Fields.Item("insSubAbsCls_id").Value%>','S');"><%=rsEquipmentClass.Fields.Item("chvSubAbsClsName").Value%></a><br>
			<DIV ID="S<%=rsEquipmentClass.Fields.Item("insSubAbsCls_id").Value%>" style="display: none">
<%
			LastConcClass = "";
			firstrow2 = 0;
		} else {
			CurrentConcClass=rsEquipmentClass.Fields.Item("insCnctCls_id").Value;
			if (CurrentConcClass != LastConcClass) {
%>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<img src="../i/leaf.gif" align="absmiddle" ALT="Concrete Class <%=rsEquipmentClass.Fields.Item("chvCnctClsName").Value%>">
				<a class="red" href="javascript: EditClass('<%=rsEquipmentClass.Fields.Item("insCnctCls_id").Value%>','C');"><%=rsEquipmentClass.Fields.Item("chvCnctClsName").Value%></a><br>
<%
			}
			LastConcClass = CurrentConcClass;
		}
		LastSubABSClass = CurrentSubABSClass;		
	}
	LastABSClass = CurrentABSClass;
	rsEquipmentClass.MoveNext();
}
%>
</DIV>
</body>
</html>
<%
rsEquipmentClass.Close();
%>
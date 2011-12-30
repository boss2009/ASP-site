<%
if(String(Session("insStaff_id")) == "undefined") {
	if (Request.QueryString == "") {
		Response.Redirect("/../timeout.asp?page="+Request.ServerVariables("URL"));
	} else {
		Response.Redirect("/../timeout.asp?page="+Request.ServerVariables("URL")+"&"+Request.QueryString);
	}
}
%>

<%
if(String(Session("insStaff_id")) == "undefined") {
	if (Request.QueryString=="") {
		Response.redirect("/../timeout.asp?page="+Request.ServerVariables("URL"));
	} else {
		Response.redirect("/../timeout.asp?page="+Request.ServerVariables("URL")+"&"+Request.QueryString);
	}
}
if((Session("MM_UserAuthorization")) < 5) {
	Response.redirect("/../dologin.asp");
}
%>
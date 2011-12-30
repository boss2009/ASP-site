<!--------------------------------------------------------------------------
* File Name: asplogout.asp
* Title: 
* Main SP: 
* Description: Ends the session.
* Author: D. T. Chan
--------------------------------------------------------------------------->
<%
Session.Abandon
Response.Redirect "asplogin.asp"
%>
<!--------------------------------------------------------------------------
* File Name: asplogout.asp
* Title: 
* Main SP: 
* Description: Ends the session.
* Author: T.H
--------------------------------------------------------------------------->
<%
Session.Abandon
Response.Redirect "asplogin.asp"
%>
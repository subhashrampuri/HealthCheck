
<!--#include file="./includes/connect.asp"-->
<%
	Session.Timeout=40
	act_ID = Request.Form("hdTime")
	Text = Request.Form("hdStatus")
	
	IF Text = "Active" then
		sDeactivate = "Update tblActivityTime Set act_IsActive=0 where act_ID='"& act_ID &"'"
		ConnFMG.Execute(sDeactivate)
	else
		sActivate = "Update tblActivityTime Set act_IsActive=1 where act_ID='"& act_ID &"'"
		ConnFMG.Execute(sActivate)
	end if	
%>	
<!--#include file="./includes/connect.asp"-->
<%
	Response.Redirect "manage_activitytime.asp"
%>

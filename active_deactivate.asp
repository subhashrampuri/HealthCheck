
<!--#include file="./includes/connect.asp"-->
<%
  	Session.Timeout=40
	EmpID = Request.Form("hdEmployeeID")
	Text = Request.Form("hdStatus")
	Location = Request.Form("hdLocation")
	
	Response.write EmpID & "~~~~" & Text
	IF Text = "Active" then
		sDeactivate = "Update tblUsers Set isActive=0 where usr_EmpID='"& EmpID &"' and loc_ShortCode = '"& Location &"'"
		ConnFMG.Execute(sDeactivate)
	else
		sActivate = "Update tblUsers Set isActive=1 where usr_EmpID='"& EmpID &"' and loc_ShortCode = '"& Location &"'"
		ConnFMG.Execute(sActivate)
	end if	
%>	
<!--#include file="./includes/connect.asp"-->
<%
	Response.Redirect "manage_admin_users.asp"
%>

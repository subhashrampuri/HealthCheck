
<!--#include file="./includes/connect.asp"-->
<%
	Session.Timeout=40
	Role =  Replace(Server.HTMLEncode(request.form("txtRole")),"'","''")
	Active = request.form("isActive")
	if Active <> "" then
		sActive = 1
	else
		sActive = 0 
	end if
	rol_id = Replace(Server.HTMLEncode(Request.form("hdrol_ID")),"'","''")	

	lclstr_ActionType = REPLACE(TRIM(Request.Form("butSubmit")),"'","")
	lclstr_redirect = "manage_add_roles.asp?msgno="

	if UCASE(lclstr_ActionType) = UCASE("Add") then
	
		MySQL = "Select rol_Name from tblRoles where rol_Name = '"& Trim(Role) &"'"
		Set lclobj_Rec = ConnFMG.Execute(MySQL)
		
		if lclobj_Rec.Eof then
			 Sql = "Insert into tblRoles values ('" & Role & "'," & sActive & ")"
			 ConnFMG.Execute(Sql)
			 lclstr_redirect = lclstr_redirect & "1"
		else
			lclstr_redirect = lclstr_redirect & "2&actionType="&lclstr_ActionType
			'&"&rol_name="&Role 
		end if	
		set lclobj_Rec = NOTHING

	elseif UCASE(lclstr_ActionType) = UCASE("Update") then
		' UPDATE Roles INFORMATION
		lclstr_MySQL = "Update tblRoles Set rol_Name='"& Trim(Role) &"',IsActive="&sActive&" where rol_ID='"& rol_id &"'"
		'response.write lclstr_MySQL
		ConnFMG.Execute(lclstr_MySQL)
		lclstr_redirect = lclstr_redirect & "3"
	else
		Response.redirect "manage_add_roles.asp"	
	end if
%>	
<!--#include file="./includes/connect.asp"-->
<%
	Response.Redirect lclstr_redirect
%>

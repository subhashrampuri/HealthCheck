
<!--#include file="./includes/connect.asp"-->
<%
	Session.Timeout=40
	Area =  Replace(Server.HTMLEncode(request.form("txtArea")),"'","''")
	Active = request.form("isActive")
	if Active <> "" then
		sActive = 1
	else
		sActive = 0
	end if
	area_id = Replace(Server.HTMLEncode(Request.form("hdarea_ID")),"'","''")

	lclstr_ActionType = REPLACE(TRIM(Request.Form("butSubmit")),"'","")
	lclstr_redirect = "manage_add_areas.asp?msgno="

	if UCASE(lclstr_ActionType) = UCASE("Add") then

		MySQL = "Select area_Name from tblAreas where area_Name = '"& Trim(Area) &"'"
		Set lclobj_Rec = ConnFMG.Execute(MySQL)

		if lclobj_Rec.Eof then
			 Sql = "Insert into tblAreas values ('" & Area & "'," & sActive & ")"
			 ConnFMG.Execute(Sql)
			 lclstr_redirect = lclstr_redirect & "1"
		else
			lclstr_redirect = lclstr_redirect & "2&actionType="&lclstr_ActionType
			'&"&area_name="&Area
		end if
		set lclobj_Rec = NOTHING

	elseif UCASE(lclstr_ActionType) = UCASE("Update") then
		' UPDATE Area INFORMATION
		lclstr_MySQL = "Update tblAreas Set area_Name='"& Area &"',IsActive="&sActive&" where area_ID='"& area_id &"'"
		'response.write lclstr_MySQL
		ConnFMG.Execute(lclstr_MySQL)
		lclstr_redirect = lclstr_redirect & "3"
	else
		Response.redirect "manage_add_areas.asp"
	end if
%>
<!--#include file="./includes/connect.asp"-->
<%
	Response.Redirect lclstr_redirect
%>

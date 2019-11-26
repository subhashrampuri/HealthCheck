
<!--#include file="./includes/connect.asp"-->
<%
	Session.Timeout=40
	EmpID =Replace(Server.HTMLEncode(request.form("txtEmpID")),"'","''")
	Emp =  Replace(Server.HTMLEncode(request.form("txtEmp")),"'","''")
	Email =Replace(Server.HTMLEncode(request.form("txtEmail")),"'","''")
	Active = request.form("isActive")
	if Active <> "" then
		sActive = 1
	else
		sActive = 0
	end if
	alert_id = Replace(Server.HTMLEncode(Request.form("hdalt_ID")),"'","''")

	lclstr_ActionType = REPLACE(TRIM(Request.Form("butSubmit")),"'","")
	lclstr_redirect = "manage_add_alertusers.asp?msgno="

	if UCASE(lclstr_ActionType) = UCASE("Add") then

		MySQL = "Select alt_Empid from tblAlertUsers where alt_Empid = '"& Trim(EmpID) &"'"
		Set lclobj_Rec = ConnFMG.Execute(MySQL)

		if lclobj_Rec.Eof then
			 Sql = "Insert into tblAlertUsers values (" & EmpID & ",'" & Emp & "','" & Email &"'," & sActive & ")"

			 ConnFMG.Execute(Sql)
			 lclstr_redirect = lclstr_redirect & "1"
		else
			lclstr_redirect = lclstr_redirect & "2&actionType="&lclstr_ActionType
			'&"&area_name="&Area
		end if
		set lclobj_Rec = NOTHING

	elseif UCASE(lclstr_ActionType) = UCASE("Update") then
		' UPDATE Area INFORMATION
		lclstr_MySQL = "Update tblAlertUsers Set alt_Empid="& EmpID &",alt_Name='"& Emp &"',alt_Email='" & Email & "',IsActive="&sActive&" where alt_ID="& alert_id &""
		'response.write lclstr_MySQL
		ConnFMG.Execute(lclstr_MySQL)
		lclstr_redirect = lclstr_redirect & "3"
	else
		Response.Redirect "manage_add_alertusers.asp"
	end if
%>
<!--#include file="./includes/connect.asp"-->
<%
	Response.Redirect lclstr_redirect
%>

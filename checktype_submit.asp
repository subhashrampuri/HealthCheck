
<!--#include file="./includes/connect.asp"-->
<%
	Session.Timeout=40
	Area =  Replace(Server.HTMLEncode(request.form("cboArea")),"'","''")
	checktype =  Replace(Server.HTMLEncode(request.form("txtChecktype")),"'","''")
	Active = request.form("isActive")
	if Active <> "" then
		sActive = 1
	else
		sActive = 0 
	end if
	'response.write Area & "<br>" & checktype & "<br>" & Active
	'response.end
	
	chktyp_id = Replace(Server.HTMLEncode(Request.form("hdchktyp_ID")),"'","''")	

	lclstr_ActionType = REPLACE(TRIM(Request.Form("butSubmit")),"'","")
	lclstr_redirect = "manage_admin_checktypes.asp?msgno="

	if UCASE(lclstr_ActionType) = UCASE("Add") then
	
		MySQL = "Select chktyp_Name from tblChecktypes where chktyp_Name = '"& Trim(checktype) &"' and area_ID = "& Trim(Area) &" "
		Set lclobj_Rec = ConnFMG.Execute(MySQL)
		
		if lclobj_Rec.Eof then
			 Sql = "Insert into tblChecktypes values (" & Area & ",'" & checktype & "'," & sActive & ")"
			 ConnFMG.Execute(Sql)
			 lclstr_redirect = lclstr_redirect & "1"
		else
			lclstr_redirect = lclstr_redirect & "2&actionType="&lclstr_ActionType
			'&"&area_name="&Area 
		end if	
		set lclobj_Rec = NOTHING

	elseif UCASE(lclstr_ActionType) = UCASE("Update") then
		' UPDATE Area INFORMATION
		lclstr_MySQL = "Update tblChecktypes Set area_ID = "& Area & ",chktyp_Name='"& checktype &"',IsActive="&sActive&" where chktyp_ID='"& chktyp_id &"'"
		'response.write lclstr_MySQL
		ConnFMG.Execute(lclstr_MySQL)
		lclstr_redirect = lclstr_redirect & "3"
	else
		Response.Redirect "manage_admin_checktypes.asp"	
	end if
%>	
<!--#include file="./includes/connect.asp"-->
<%
	Response.Redirect lclstr_redirect
%>

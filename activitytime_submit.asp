
<!--#include file="./includes/connect.asp"-->
<%
	Session.Timeout=40
	sTime =  Replace(Server.HTMLEncode(Trim(request.form("cboTime"))),"'","''")
	Active = request.form("isActive")
	sDisplay = Replace(Server.HTMLEncode(Trim(request.form("txtDisplay"))),"'","''")

	if Active <> "" then
		sActive = 1
	else
		sActive = 0 
	end if
	'response.write sTime & "<br>" & Active
	'response.end
	
	chktyp_id = Replace(Server.HTMLEncode(Request.form("hdact_ID")),"'","''")	

	lclstr_ActionType = REPLACE(TRIM(Request.Form("butSubmit")),"'","")
	lclstr_redirect = "manage_activitytime.asp?msgno="

	if UCASE(lclstr_ActionType) = UCASE("Add") then
	
		MySQL = "Select act_Time from tblActivityTime where act_Time = '"& Trim(sTime) &"'"
		Set lclobj_Rec = ConnFMG.Execute(MySQL)
		
		if lclobj_Rec.Eof then
			 Sql = "Insert into tblActivityTime values ('"& Trim(sTime)&"'," & sActive & ",'"& sDisplay &"')"
			 ConnFMG.Execute(Sql)
			 lclstr_redirect = lclstr_redirect & "1"
		else
			lclstr_redirect = lclstr_redirect & "2&actionType="&lclstr_ActionType
			'&"&act_Time="&sTime 
		end if	
		set lclobj_Rec = NOTHING
	else
		Response.redirect "manage_activitytime.asp"
	end if
%>	
<!--#include file="./includes/connect.asp"-->
<%
	Response.Redirect lclstr_redirect
%>

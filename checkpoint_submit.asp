
<!--#include file="./includes/connect.asp"-->
<%
	Session.Timeout=40
	Area =  request.form("cboArea")
	checktype =  request.form("cboChktype")
	lclstr_checkpoint = Replace(Request.form("txtChkpoint"),"'","''")
	Active = request.form("IsActive")
	if Active <> "" then
		sActive = 1
	else
		sActive = 0
	end if

	chkpnt_id = Replace(Server.HTMLEncode(Request.form("hdchkpnt_ID")),"'","''")
	lclstr_ActionType = REPLACE(TRIM(Request.Form("butSubmit")),"'","")
	lclstr_redirect = "manage_admin_checkpoints.asp?msgno="

	if UCASE(lclstr_ActionType) = UCASE("Add") then

		MySQL = "Select chkpnt_name from tblCheckpoint where area_ID ="& Area &" and chktyp_ID=" & checktype & " and chkpnt_name='" & trim(lclstr_checkpoint) & "'"
		Set lclobj_Rec = ConnFMG.Execute(MySQL)

		if lclobj_Rec.Eof then
			 Sql = "Insert into tblCheckpoint values (" & Area & "," & checktype & ",'" & lclstr_checkpoint & "'," & sActive & ")"
			 ConnFMG.Execute(Sql)

			'	if  request.form("txtaltusers") <> "" then
			'		lclint_altuserid = request.form("txtaltusers")
			'		vArray = Split(lclint_altuserid,",")

			'		sQuery = " Select Max(chkpnt_ID) as chkpnt_ID from tblCheckpoint "
			'		Set objRs = ConnFMG.Execute(sQuery)

			'			if objRs.Eof = false then
			'				max_chkpnt_ID = objRs("chkpnt_ID")
			'			end if

			'			for i = 0 to UBound(vArray)

			'				sInsert = "Insert Into tblCheckpoint_Map_Alertuser Values("& max_chkpnt_ID &",'"& Trim(vArray(i)) &"')"
			'				ConnFMG.Execute(sInsert)
							'response.write vArray(i) & "<br>"
			'			Next
			'	end if

			 lclstr_redirect = lclstr_redirect & "1"
		else
			lclstr_redirect = lclstr_redirect & "2&actionType="&lclstr_ActionType
		end if
		set lclobj_Rec = NOTHING

	elseif UCASE(lclstr_ActionType) = UCASE("Update") then

		lclstr_MySQL = "Update tblCheckpoint Set area_ID = "& Area & ",chktyp_ID="& checktype &",chkpnt_name='" & lclstr_checkpoint & "',IsActive="&sActive&" where chkpnt_ID="& chkpnt_id  &""
		ConnFMG.Execute(lclstr_MySQL)

			'if  request.form("txtaltusers") <> "" then
			'	lclint_altuserid = request.form("txtaltusers")
			'	vArray = Split(lclint_altuserid,",")

			'	sDelete = " Delete from tblCheckpoint_Map_AlertUser where chkpnt_ID= "& chkpnt_id  &""
			'	ConnFMG.Execute(sDelete)

			'		for i = 0 to UBound(vArray)

			'			sInsert = "Insert Into tblCheckpoint_Map_Alertuser Values("& chkpnt_id &",'"& Trim(vArray(i)) &"')"
			'			ConnFMG.Execute(sInsert)

			'		Next
			'end if

		lclstr_redirect = lclstr_redirect & "3"
	else
		Response.Redirect "manage_admin_checkpoints.asp"
	end if
%>
<!--#include file="./includes/connect.asp"-->
<%
	Response.Redirect lclstr_redirect
%>

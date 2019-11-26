<%@LANGUAGE="VBSCRIPT"%>
	<% if  not (session("sms_validated")) = "True" then
		session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
		Response.Redirect "/default.asp"
		end if
	%>

<!--#include file="./includes/connect.asp"-->
	<%
		Session.Timeout=40
		Server.ScriptTimeout = 300
		Set adoRs = Server.CreateObject("ADODB.RecordSet")
		Set adocmd  = Server.createobject("ADODB.Command")
		Emp_ID = Request.form("hdEmp_ID")
		sFilter = Request.form("hdFilter")
		Location = Request.form("hdLocation")	
		Slot = Request.form("hdSlot")
		EntryBy = Request.Form("hdEntryBy")
		Summary = Replace(Server.HTMLEncode(Request.Form("txtSummary")),"'","''")
	'	Response.write Emp_ID & ":" & sFilter & ":" & Location & ":" & Slot & ":" & EntryBy
		
		
		ChkMaster = "select tdm_ID,act_Time,tdm_Date from tblTransactionDetailsMaster where DateDiff(DD,tblTransactionDetailsMaster.tdm_Date,'"& sFilter &"') = 0 " & _ 
			" and usr_EmpID = '"& Trim(Emp_ID)&"' and loc_ShortCode = '"& trim(Location) &"' and act_Time = '"& Slot &"' "
		Set objRs = ConnFMG.Execute(ChkMaster)		

		if objRs.Eof = false then
			Del_ID = objRs("tdm_ID")
			Del_Date = objRs("tdm_Date")

				DeleteMaster = "Delete from tblTransactionDetailsMaster where tdm_ID = "&Del_ID&" "
				DeleteTrans	= "Delete from tblTransactionDetails where tdm_ID = "&Del_ID&" "
				
				ConnFMG.Execute(DeleteMaster)
				ConnFMG.Execute(DeleteTrans)

				if request.form("txtAttachment") <> "" then
					attachment = request.form("txtAttachment")
					sql = "sp_itblTransactionDetailsMaster '"& Trim(Emp_ID) &"','"& attachment &"','"& slot &"','"& Del_Date &"','"& Trim(Location) &"','"& Trim(Summary) & "','"& Trim(EntryBy)&"','"& Trim(Session("sms_EmplNbr")) &"','"&now()&"' "
				else
					attachment = NULL
					sql = "sp_itblTransactionDetailsMaster '"& Trim(Emp_ID) &"',null,'"& slot &"','"& Del_Date &"','"& Trim(Location) &"','"& Trim(Summary) & "','"& Trim(EntryBy)&"','"& Trim(Session("sms_EmplNbr")) &"','"&now()&"' "
				end if
		
				adoRs.Open sql,ConnFMG,3,3
				ReqID=adoRs(0)
				adors.Close
				Set adors = Nothing
		
				Set objLoop = Server.CreateObject("ADODB.RecordSet")
				sLoop = "Select area_ID,chktyp_ID,chkpnt_ID,loc_ShortCode from tblTransaction where usr_EmpID= '"& Trim(Emp_ID)&"' " & _ 
					" and loc_shortcode = '"& trim(Location) &"' order by area_ID,chktyp_ID,chkpnt_ID " 
				objLoop.Open sLoop,ConnFMG,3,3
				While Not objLoop.EOF
				ID =  objLoop("area_ID")&"_"&objLoop("chktyp_ID")&"_"&objLoop("chkpnt_ID")
		
				sTransData=""
				if Trim(request.form("rdo_"&ID)) <> ""  then
					if Trim(request.form("txtcomment_"&ID)) <> "" then 
						sTransData = " Insert into tblTransactionDetails values ('"& ReqID &"', '"& trim(Emp_ID)&"',"& objLoop("area_ID") &","& objLoop("chktyp_ID") &","& objLoop("chkpnt_ID") &",'"& Replace(Server.HTMLEncode(Request.form("txtcomment_"&ID)),"'","''") &"','"&now()&"', "& Trim(request.form("rdo_"&ID)) &",'"& Trim(location) &"')"
					else
						sTransData = " Insert into tblTransactionDetails values ('"& ReqID &"', '"& trim(Emp_ID)&"',"& objLoop("area_ID") &","& objLoop("chktyp_ID") &","& objLoop("chkpnt_ID") &",null,'"&now()&"', "& Trim(request.form("rdo_"&ID)) &",'"& Trim(location) &"')"
					end if
					ConnFMG.Execute(sTransData)
				end if
				
				objLoop.MoveNext
				Wend
else

				if request.form("txtAttachment") <> "" then
					attachment = request.form("txtAttachment")
					sql = "sp_itblTransactionDetailsMaster '"& Trim(Emp_ID) &"','"& attachment &"','"& slot &"','"& now() &"','"& Trim(Location) &"','"& Trim(Summary) & "','"& Trim(EntryBy)&"','"& Trim(Session("sms_EmplNbr")) &"','"&now()&"' "
				else
					attachment = NULL
					sql = "sp_itblTransactionDetailsMaster '"& Trim(Emp_ID) &"',null,'"& slot &"','"& now() &"','"& Trim(Location) &"','"& Trim(Summary) & "','"& Trim(EntryBy)&"','"& Trim(Session("sms_EmplNbr")) &"','"&now()&"' "
				end if
		
				adoRs.Open sql,ConnFMG,3,3
				ReqID=adoRs(0)
				adors.Close
				Set adors = Nothing
		
				Set objLoop = Server.CreateObject("ADODB.RecordSet")
				sLoop = "Select area_ID,chktyp_ID,chkpnt_ID,loc_ShortCode from tblTransaction where usr_EmpID= '"& Trim(Emp_ID)&"' " & _ 
					" and loc_shortcode = '"& trim(Location) &"' order by area_ID,chktyp_ID,chkpnt_ID " 
				objLoop.Open sLoop,ConnFMG,3,3
				While Not objLoop.EOF
				ID =  objLoop("area_ID")&"_"&objLoop("chktyp_ID")&"_"&objLoop("chkpnt_ID")
		
				sTransData=""
				if Trim(request.form("rdo_"&ID)) <> ""  then
					if Trim(request.form("txtcomment_"&ID)) <> "" then 
						sTransData = " Insert into tblTransactionDetails values ('"& ReqID &"', '"& trim(Emp_ID)&"',"& objLoop("area_ID") &","& objLoop("chktyp_ID") &","& objLoop("chkpnt_ID") &",'"& Replace(Server.HTMLEncode(Request.form("txtcomment_"&ID)),"'","''") &"','"&now()&"', "& Trim(request.form("rdo_"&ID)) &",'"& Trim(location) &"')"
					else
						sTransData = " Insert into tblTransactionDetails values ('"& ReqID &"', '"& trim(Emp_ID)&"',"& objLoop("area_ID") &","& objLoop("chktyp_ID") &","& objLoop("chkpnt_ID") &",null,'"&now()&"', "& Trim(request.form("rdo_"&ID)) &",'"& Trim(location) &"')"
					end if
					ConnFMG.Execute(sTransData)
				end if
				
				objLoop.MoveNext
				Wend

end if
  %>
	  
<form name="frmSlot">
	<input type="hidden" name="hdSlot" value="<%=slot%>">
</form>
<script language="javascript">
	document.frmSlot.method = "Post";
	document.frmSlot.action ="helpdesk_activitytime_ack.asp"
	document.frmSlot.submit()
</script>

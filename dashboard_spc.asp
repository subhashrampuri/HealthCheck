<%@LANGUAGE="VBSCRIPT"%>
<% if  not (session("sms_validated")) = "True" then
	session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
	Response.Redirect "/default.asp"
 	end if
		Session.Timeout=40
	Server.ScriptTimeout = 300
%>
<!--#include file="./includes/connect.asp"-->
<%
	GetSlot = "Select top 1 act_Time from tblActivityTime where act_IsActive = 1"
	Set RsSlot = ConnFMG.Execute(GetSlot)
	if RsSlot.EOF = false then
		CurSlot = RsSlot("act_Time")
	end if
	RsSlot.Close
%>
<%
if request.Form("cboFromDD") <> "" then
	tDay = request.Form("cboFromDD")
else
	tDay = day(now())
end if
if request.Form("cboFromMM") <> "" then
	tMonth = request.Form("cboFromMM")
else
	tMonth = month(now())
end if
if request.Form("cboFromYY") <> "" then
	tYear = request.Form("cboFromYY")
else
	tYear = Year(now())
end if
if request.form("cboSlot") <> "" then
	tSlot = request.form("cboSlot")
else
	tSlot = CurSlot
end if
	tDate = tMonth & "/" & tDay & "/" & tYear
	'response.write tDate & "<br>"
	'response.write tSlot & "<br>"
	
	'GetRole = "Select rol_ID,loc_ShortCode from tblUsers where usr_EmpID = '"& Trim(Session("sms_EmplNbr")) &"'"
	'set objRole = ConnFMG.Execute(GetRole)
	'if objRole.EOF = false then
	'if objRole("rol_ID") = "4" then
	'	iExist = 1
	'	Loc_ID = Trim(objRole("loc_ShortCode"))
	'	GetHD = "Select usr_EmpID from tblUsers where usr_EmpID <> '"& Trim(Session("sms_EmplNbr")) &"' and loc_ShortCode = '"& Trim(objRole("loc_ShortCode")) &"'" 
	'	set objGetHD = ConnFMG.Execute(GetHD)
	'	if objGetHD.EOF = false then
	'		Emp_ID = Trim(objGetHD("usr_EmpID"))	
	'	end if
	'else
	'	iExist = 0
		Emp_ID = Trim(Session("sms_EmplNbr"))	
	'end if
	'end if
	'objRole.Close
	'Set objRole = nothing

	
%>
<html>
<head>
<title>Welcome to our Intranet | Health Check</title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
<Script language=JavaScript src="./includes/validate.js" type=text/javascript></SCRIPT>
<script language="javascript">
	function Validate_Form()
	{
		var frmobj = document.frmDashboard;
		if (frmobj.cboFromDD.value=="0")
		{
			alert("Please select day")
			frmobj.cboFromDD.focus()
			return false;
		}
		if (frmobj.cboFromMM.value=="0")
		{
			alert("Please select month")
			frmobj.cboFromMM.focus()
			return false;
		}
		if (frmobj.cboFromYY.value=="0")
		{
			alert("Please select year")
			frmobj.cboFromYY.focus()
			return false;
		}
		if (frmobj.cboSlot.value=="0")
		{
			alert("Please select slot")
			frmobj.cboSlot.focus()
			return false;
		}
//		isValidDate
		var sdt = frmobj.cboFromDD.value;
		var smt = frmobj.cboFromMM.value;
		var syr = frmobj.cboFromYY.value;

		sdate = (sdt+"/"+smt+"/"+syr);
	/*	if (!isValidDate(sdate))
		{
			alert("Please select valid date");
			frmobj.cboFromDD.focus();
			return false;
		}
	*/
		return true

	}
	function Edit_SPC(LOC,EMPID)
	{
	//alert(document.frmSPCEdit.hdEmp_ID.value);
	//alert(document.frmSPCEdit.hdSlot.value);
	//alert(document.frmSPCEdit.hdFilter.value);
	//alert(document.frmSPCEdit.hdLocation.value);				
	 document.frmSPCEdit.hdLocation.value = LOC;
	 document.frmSPCEdit.hdEmp_ID.value = EMPID;
	 document.frmSPCEdit.method="Post";
	 document.frmSPCEdit.action="helpdesk_activity_edit.asp"
	 document.frmSPCEdit.submit()	 	 
	}
</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#ffffff">
  <table cellpadding="0" cellspacing="0" border="0" height="100%" bgcolor="#ffffff">
    <tr>

      <td valign="top" bgcolor="#F0F8FE" width="59"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
		<td valign="top" bgcolor="#037DDB" width="1" ><img src="./images/blank.gif" border="0" width="1" height="1"></td>
		<!-- Middle Part Start Here -->

      <td valign="top" bgcolor="#FFFFFF" width="879">
        <table cellpadding="0" cellspacing="0" border="0" width="100%"  vAlign="top">
               <tr>
			   		<td valign="top">
						<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff"  vAlign="top">
						    <tr><td valign="top" height="1" colspan="5"><img src="images/blank.gif" border="0" width="1" height="1"></td></tr>
							<tr>
								<td valign="top" width="5"><img src="images/blank.gif" border="0" width="15" height="1"></td>
								<td valign="top" width="80%" align="lefet"><img src="images/logo.jpg" border="0"></td>

                  <td valign="bottom" align="center" width="5%">&nbsp;</td>

                  <td valign="bottom" align="center" width="5%">&nbsp;</td>

                  <td valign="bottom" align="center" width="5%">&nbsp;</td>

                  <td valign="bottom" align="center" width="5%">&nbsp;</td>
								<td valign="top" width="5"><img src="images/blank.gif" border="0" width="1" height="1"></td>
							</tr>
							<tr><td valign="top" height="4" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td></tr>
							<tr><td valign="top" bgcolor="#4FA3E4" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td></tr>
							<tr><td valign="top" height="2" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td></tr>
							<tr>
								<td valign="top" colspan="6">
									<table cellpadding="0" cellspacing="0" border="0" width="100%">
										<tr>
											<td valign="top"><img src="images/leftchk.gif" border="0"></td>
											<td valign="top" width="1"><img src="images/blank.gif" border="0" width="1" height="1"></td>
											<td valign="top" bgcolor="#3F89C3" width="70%" height="31"><img src="images/blank.gif" border="0" width="1" height="1"></td>
											<td valign="top" width="1" ><img src="images/blank.gif" border="0" width="1" height="1"></td>
											<td valign="top"><img src="images/rightchk.gif" border="0" ></td>


                      <td valign="middle" bgcolor="#3F89C3" width="20%" height="31" class="txtfnt">&nbsp;&nbsp;<b>User
                        : <%=Session("sms_name") & " " & "(" & Session("sms_EmplNbr") & ")" %></b></td>
										</tr>
						            </table>
								</td>
							</tr>
							<tr><td valign="top" height="2" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td></tr>
							<tr><td valign="top" bgcolor="#4FA3E4" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td></tr>
							<tr><td valign="top" height="2" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td></tr>
						</table>
					</td>
				</tr>
			   <tr>
					<td valign="top">
						<table cellpadding="0" cellspacing="0" border="0" width="100%" valign="top">
                            <tr>
								<td valign="top" width="100%">


                  <table cellpadding="0" cellspacing="0" border="0" width="100%" vAlign="Top">
				  <%if Trim(Session("sms_EmplNbr")) = "1247" then%>
				  <!--#include file="./includes/navigation.asp"-->
				  <%elseif Role_ID = "1" then %>
				  <!--#include file="./includes/navigation.asp"-->
				  <% elseif Role_ID = "2" then %>							
				  <!--#include file="./includes/navigation-fmg.asp"-->
				  <% elseif Role_ID = "3" or Role_ID = "4" then %>
				  <!--#include file="./includes/navigation-hd.asp"-->
				  <%else%>
					<% Response.redirect "./index.asp" %>
				  <%end if%>

                    <tr>
                      <td valign="top" colspan="3"><br>
                      </td>
                    </tr>
                    <tr>
                      <td valign="top"  colspan="3">
                        <table cellpadding="0" cellspacing="0" border="0" width="100%">
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top" align="right">&nbsp;</td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top"> 
                              <form name="frmDashboard" method="post" action="dashboard_spc.asp" onSubmit="return Validate_Form();">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <tr align="left"> 
                                    <td colspan="2" class="txtHeader" bgcolor="#FFFFFF"><b>Filter</b></td>
                                  </tr>
                                  <tr class="Content"> 
                                    <td align="right" bgcolor="#B4DAF8" width="35%">Date 
                                      : </td>
                                    <td bgcolor="#B4DAF8" align="left"> 
                                      <select name="cboFromDD"  size="1" class="Standard">
                                        <option value="0" selected>Day</option>
                                        <% for i = 1 to 31
										if cInt(tDay) = i  then
										%>
                                        <option value="<%=i%>" selected><%=i%></option>
                                        <%else%>
                                        <option value="<%=i%>"><%=i%></option>
                                        <%end if%>
                                        <%next%>
                                      </select>
                                      <select name="cboFromMM"  size="1"  class="Standard">
                                        <option value="0" selected>Month</option>
                                        <% if cInt(tMonth) = 1  then%>
                                        <option value="1" selected>January</option>
                                        <%else%>
                                        <option value="1">January</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 2  then%>
                                        <option value="2" selected>February</option>
                                        <%else%>
                                        <option value="2">February</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 3  then%>
                                        <option value="3" selected>March</option>
                                        <%else%>
                                        <option value="3">March</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 4  then%>
                                        <option value="4" selected>April</option>
                                        <%else%>
                                        <option value="4">April</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 5  then%>
                                        <option value="5" selected>May</option>
                                        <%else%>
                                        <option value="5">May</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 6  then%>
                                        <option value="6" selected>June</option>
                                        <%else%>
                                        <option value="6">June</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 7  then%>
                                        <option value="7" selected>July</option>
                                        <%else%>
                                        <option value="7">July</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 8  then%>
                                        <option value="8" selected>August</option>
                                        <%else%>
                                        <option value="8">August</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 9  then%>
                                        <option value="9" selected>Sepember</option>
                                        <%else%>
                                        <option value="9">Sepember</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 10  then%>
                                        <option value="10" selected>October</option>
                                        <%else%>
                                        <option value="10">October</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 11  then%>
                                        <option value="11" selected>November</option>
                                        <%else%>
                                        <option value="11">November</option>
                                        <%end if%>
                                        <% if cInt(tMonth) = 12  then%>
                                        <option value="12" selected>December</option>
                                        <%else%>
                                        <option value="12">December</option>
                                        <%end if%>
                                      </select>
                                      <select name="cboFromYY"  size="1" class="Standard">
                                        <option value="0" selected>Year</option>
                                        <%
											For k = 2007 to 2010

											if Cint(tYear) = k  then
										%>
                                        <option value="<%=k%>" selected><%=k%></option>
                                        <%else%>
                                        <option value="<%=k%>"><%=k%></option>
                                        <%end if%>
                                        <%next%>
                                      </select>
                                      <%
									  	Set oSlot = Server.CreateObject("ADODB.RecordSet")
										sSlotTime =  "Select act_Time from tblActivityTime where act_IsActive = 1"
										oSlot.Open sSlotTime,ConnFMG,3,3
									  %>
                                      <select name="cboSlot">
                                        <!--<option value="0" selected>Slot</option>-->
                                        <%
											While Not oSlot.EOF
											if trim(cStr(tSlot)) = trim(oSlot("act_Time")) then
												sSelect = "Selected"
											else
												sSelect = ""
											end if
										%>
                                        <option value="<%=oSlot("act_Time")%>" <%=sSelect%>><%=FormatDateTime(oSlot("act_Time"),3)%></option>
                                        <%
										  	oSlot.Movenext
											Wend
											oSlot.Close
										  %>
                                      </select>
                                    </td>
                                  </tr>
                                  <tr class="Content" bgcolor="#B4DAF8"> 
                                    <td colspan="2">&nbsp;</td>
                                  </tr>
                                  <tr class="Content"> 
                                    <td bgcolor="#B4DAF8">&nbsp;</td>
                                    <td bgcolor="#B4DAF8"> 
                                      <input type="submit" name="Submit" value="Filter">
                                    </td>
                                  </tr>
                                </table>
                              </form>
                            </td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top" class="txtHeader"> 
                              <table width="100%" border="0" cellspacing="1" cellpadding="3">
                                <tr class="txtHeader"> 
                                  <td width="40%" align="right"><b>Legends : </b></td>
                                  <td width="4%" bgcolor="#339933">&nbsp;</td>
                                  <td>Successful</td>
                                  <td width="4%" bgcolor="#FF0000">&nbsp;</td>
                                  <td>Failure</td>
                                  <td width="4%" bgcolor="#4682b4">&nbsp;</td>
                                  <td>Not Updated</td>
                                  <td width="4%" bgcolor="#999999">&nbsp;</td>
                                  <td>Not Applicable</td>
                                </tr>
                              </table>
                            </td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top" class="txtHeader" align="left"> 
                              <%'=msg%>
                            </td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="middle" class="txtHeader">&nbsp;</td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
						 <%
						 if tDay <> "" then
						  if iExist = 1 then 
						 %>	
                          <!--<tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="middle" class="txtHeader"> 
                              <table width="100%" border="0" cellspacing="1" cellpadding="3">
                                <tr class="Content"> 
                                  <td width="53%" bgcolor="#9FC4E1" align="left"><b>Would 
                                    you like to edit activity details?</b></td>
                                  <td bgcolor="#B4DAF8" align="right"><b><a href="javascript: Edit_SPC()" class="Content">Edit</a></b></td>
                                </tr>
                              </table>
                            </td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>-->
						<% 
							end if 
						  end if
						%>	
                          <tr>
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top" class="txtHeader">&nbsp;</td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top" class="txtHeader"> 
                              <%
						if tDay <> "" then
							if IsDate(tDate) = "True" then
								tFilterDate=CDate(tDate)
						'	else
						'		response.write  "<font color=red size=2>" & "InValid date Submitted" & "</font>"
						'		Response.end
							end if
						  %>
                              <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                <tr> 
                                  <td class="txtHeader" colspan="14"><b>Health 
                                    Check Dashboard</b></td>
                                </tr>
                                <%
								set objRs = server.CreateObject("ADODB.RecordSet")
								sQuery = " Select a.area_ID,a.area_name, b.chktyp_id, b.chktyp_name, c.chkpnt_id, c.chkpnt_name " & _
									" from tblAreas as a, tblChecktypes as b, tblCheckpoint as c where a.isActive=1 and b.isActive=1 and c.isActive=1 " & _
									" and b.area_ID=a.area_ID and c.chktyp_id=b.chktyp_id order by a.area_name,b.chktyp_name, c.chkpnt_name "
								objRs.Open sQuery,ConnFMG,3,3
								area = ""
								checktype = ""
								checkpoint = ""

							%>
                                <tr class="Content" valign="middle" > 
                                  <td bgcolor="#9FC4E1" colspan="3" ><b>Areas</b></td>
                                  <td bgcolor="#9FC4E1" ><b>Check Types</b></td>
                                  <td bgcolor="#9FC4E1" ><b>Check Points</b></td>
                                  <td bgcolor="#9FC4E1" >&nbsp;</td>
                                  <%
								  	Set objLoc = Server.CreateObject("ADODB.Recordset")
								 	'RLoc = "Select a.loc_ShortCode from tblUsers a inner join tblLocations b on " & _
									'	" a.loc_ShortCode = b.loc_ShortCode where a.Usr_EmpID = '"&Trim(Emp_ID)&"'"
									RLoc = "Select a.loc_ShortCode,(select usr_EmpID from tblUsers where loc_ShortCode = a.loc_ShortCode and rol_ID = 3 and isActive = 1) as HPD_ID " & _ 
										" from tblUsers a inner join tblLocations b on  a.loc_ShortCode = b.loc_ShortCode where a.Usr_EmpID = '"& Emp_ID &"' "
									objLoc.Open RLoc,ConnFMG,3,3
									
									While Not objLoc.EOF
									locCode = Left(objLoc("loc_ShortCode"),4)
									E_ID = objLoc("HPD_ID")
								  %>
                                  <td bgcolor="#9FC4E1" align="center"><b><a href="javascript: Edit_SPC('<%=locCode%>','<%=E_ID%>')" class="Content"><%=locCode%></a></b></td>
                                  <%
								  	objLoc.MoveNext
								  	Wend
								  	objLoc.Close
								  %>
                                </tr>
                                <tr class="Content" valign="top"> 
                                  <td colspan="6" bgcolor="#B4DAF8" >Attachments</td>
                                  <%
								  	Set RsLoc = Server.CreateObject("ADODB.Recordset")
									'sLocAttach = "Select a.loc_ShortCode from tblUsers a inner join tblLocations b on " & _
									'	" a.loc_ShortCode = b.loc_ShortCode where a.Usr_EmpID = '"&Trim(Emp_ID)&"'"
									sLocAttach = "Select a.loc_ShortCode,(select usr_EmpID from tblUsers where loc_ShortCode = a.loc_ShortCode and rol_ID = 3 and isActive = 1) as HPD_ID " & _ 
										" from tblUsers a inner join tblLocations b on  a.loc_ShortCode = b.loc_ShortCode where a.Usr_EmpID = '"& Emp_ID &"' "
									 
									RsLoc.Open sLocAttach,ConnFMG,3,3
									While Not RsLoc.EOF
										sUserID = RsLoc("HPD_ID")
										'GetEMPID = "select usr_EmpID from tblUsers where loc_ShortCode = '"& RsLoc("loc_ShortCode") &"' and rol_ID = 3"
										'Set RsEmp = ConnFMG.Execute(GetEMPID)
										'if RsEmp.Eof = false then
										'	HPD_ID = RsEmp("usr_EmpID")
										'end if	
										'RsEmp.Close
									  	sAttach = "Select tdm_UploadDocument from tblTransactionDetailsMaster where tdm_UploadDocument is not null and  loc_ShortCode='"& RsLoc("loc_ShortCode") &"' " & _
											" and datediff(dd,'"& tFilterDate &"',tdm_Date)=0 and usr_EmpID = '"& Trim(sUserID)&"'  and act_Time = '"& tSlot &"' "
										'response.write sAttach	
										Set adoAttach = ConnFMG.Execute(sAttach)
										if Not adoAttach.EOF then
										'	if adoAttach("tdm_UploadDocument") <> "" then
											
											if ISEMPTY(adoAttach("tdm_UploadDocument")) or ISNULL(adoAttach("tdm_UploadDocument")) then
												iAttach = 1
											else
												touch = adoAttach("tdm_UploadDocument")
												iAttach = "<a href='" & touch & "' target=_blank><img src=images/save.gif border=0></a>"
											end if
										
										end if

								  %>
                                  <td  bgcolor="#B4DAF8" align="center"><%=iAttach%></td>
                                  <%
								   iAttach = ""
								   	RsLoc.Movenext
									Wend
									RsLoc.Close
								   %>
                                </tr>
                                <%
								While Not objRs.EOF
								if area <> objRs("area_name") then
								  	area = objRs("area_name")
								%>
                                <tr class="Content" valign="middle" bgcolor="#B4DAF8"> 
                                  <td colspan="21" ><b><%=area%> </b></td>
                                </tr>
                                <%end if%>
                                <%
								if checktype <> objRs("chktyp_Name") then
								   checktype = objRs("chktyp_Name")
							  %>
                                <tr class="Content" valign="top"> 
                                  <td width="2%" bgcolor="#B4DAF8" >&nbsp;</td>
                                  <td colspan="20" bgcolor="#B4DAF8" ><b><%=checktype%></b></td>
                                </tr>
                                <%end if%>
                                <%
								if checkpoint <> objRs("chkpnt_name") then
									checkpoint = objRs("chkpnt_name")
							  %>
                                <tr class="Content" valign="top"> 
                                  <td bgcolor="#B4DAF8" colspan="2" >&nbsp;</td>
                                  <td bgcolor="#B4DAF8" colspan="4" ><%=checkpoint%></td>
                                  <%

								  	Set adoLoc = Server.CreateObject("ADODB.Recordset")
									Set adoRs = Server.CreateObject("ADODB.Recordset")
									'sLocation = "Select a.loc_ShortCode from tblUsers a inner join tblLocations b on " & _
									'	" a.loc_ShortCode = b.loc_ShortCode where a.Usr_EmpID = '"&Trim(HPD_ID)&"'"
									sLocation = "Select a.loc_ShortCode,(select usr_EmpID from tblUsers where loc_ShortCode = a.loc_ShortCode and rol_ID = 3 and isActive = 1) as HPD_ID " & _ 
										" from tblUsers a inner join tblLocations b on  a.loc_ShortCode = b.loc_ShortCode where a.Usr_EmpID = '"& Emp_ID &"' "
									
									'Response.write sLocation	
									adoLoc.Open sLocation,ConnFMG,3,3
									While Not adoLoc.EOF
									HPD_ID = adoLoc("HPD_ID")
									sDash = "select chkpnt_ID from tblTransaction where usr_EmpID = '"&Trim(HPD_ID)&"' and loc_Shortcode ='"& adoLoc("loc_ShortCode")&"' and chkpnt_ID = "& objRs("chkpnt_ID") &" "
 								    Set adoRs = ConnFMG.Execute(sDash)
										if Not adoRs.EOF then

										sMatrixData = "SELECT tblTransactionDetails.tdt_Comments, tblTransactionDetailsMaster.tdm_UploadDocument,tblTransactionDetails.tdt_Date, "& _
											" tblTransactionDetails.tdt_ActivityStatus FROM tblTransactionDetails INNER JOIN tblTransactionDetailsMaster ON " & _
											" tblTransactionDetails.tdm_ID = tblTransactionDetailsMaster.tdm_ID Where tblTransactionDetails.usr_EmpID = '"&Trim(HPD_ID)&"' " & _
											" and tblTransactionDetails.chkpnt_ID = "& objRs("chkpnt_ID") &" and datediff(dd,'"& tFilterDate &"',tblTransactionDetails.tdt_Date)=0 and " & _
											" tblTransactionDetailsMaster.act_Time = '"& tSlot &"' and tblTransactionDetails.loc_ShortCode = '"& adoLoc("loc_ShortCode")&"' "
											'Response.write 	sMatrixData
											Set adoMatrix = ConnFMG.Execute(sMatrixData)
												if Not adoMatrix.Eof then
													status = adoMatrix("tdt_ActivityStatus")
													comments = adoMatrix("tdt_Comments")
													if status = 1 then
														color = "#339933"
													elseif status = 2 then
														color = "#FF0000"
													end if
												else
													status = ""
													color="#4682b4"
												end if
										else
											status = ""
											color="#999999"
										end if

								  %>
                                  <td  bgcolor="<%=color %>" title="<%=comments%>"> 
                                    <%'=adoLoc("loc_ShortCode")%>
                                  </td>
                                  <%
								  	comments = ""
								  	adoLoc.MoveNext
									Wend
									adoLoc.Close
								  %>
                                </tr>
                                <%end if%>
                                <%
									objRs.MoveNext
									Wend
									objRs.Close
									'adoRs.Close
									Set objRs = Nothing
									Set adoRs = nothing
								%>
                                <tr class="Content" valign="top"> 
                                  <td bgcolor="#B4DAF8" colspan="21" >&nbsp;</td>
                                </tr>
                                <tr class="Content" valign="top"> 
                                  <td bgcolor="#B4DAF8" colspan="21" >&nbsp;</td>
                                </tr>
                                <tr class="Content" valign="top"> 
                                  <td colspan="21" bgcolor="#9FC4E1" >&nbsp;&nbsp;</td>
                                </tr>
                              </table>
                              <%end if%>
                            </td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">&nbsp;</td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
								</td>

							</tr>
                        </table>
					</td>
				</tr>
				<tr><td valign="top" height="20"><img src="./images/blank.gif" border="0" width="1" height="1"></td></tr>
            </table>
		</td>
		<!-- Middle Part End Here -->
		<td valign="top" bgcolor="#037DDB" width="1"><img src="./images/blank.gif" border="0" width="1" height="1"></td>

      <td valign="top" bgcolor="#F0F8FE" width="78"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
	</tr>
  </table>
  <form name="frmSPCEdit">
  <input type="hidden" name="hdFilter" value="<%=Trim(tDate)%>">
  <input type="hidden" name="hdSlot" value="<%=Trim(tSlot)%>">
  <input type="hidden" name="hdLocation" value="">  
  <input type="hidden" name="hdEmp_ID" value="">  
  </form>
</body>
</html>
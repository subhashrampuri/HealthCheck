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
	sLocation = Request.form("cboLocation")

	EmpID = Trim(Session("sms_EmplNbr"))
		Set objRs =  Server.CreateObject("ADODB.Recordset")
		sQuery = "SELECT  tblRoles.rol_Name, tblLocations.loc_Name, tblLocations.loc_ShortCode, tblUsers.usr_EmpID, tblUsers.usr_Name, tblUsers.usr_Email " & _
			" FROM  tblUsers INNER JOIN tblRoles ON tblUsers.rol_ID = tblRoles.rol_ID INNER JOIN " & _
			" tblLocations ON tblUsers.loc_ShortCode = tblLocations.loc_ShortCode WHERE (tblUsers.usr_EmpID = '"& Trim(Session("sms_EmplNbr"))&"' and tblUsers.loc_ShortCode = '"& sLocation &"')"
		objRs.Open sQuery,ConnFMG,3,3
%>
 <%
	Set adoRs = Server.CreateObject("ADODB.Recordset")
	sSlotTrack = " Select top 1 *,GetDate() as theDate from tblActivityTime where (DATEPART (hh, act_time)) > (DATEPART (hh, getdate())) and act_Isactive = 1 " & _
			" Union ALL " & _
			" Select top 1 a.*,DateAdd(DD,1,GetDate()) as theDate from tblActivityTime as a where (DATEPART (hh, a.act_time)) = (Select MIN((DATEPART (hh, a1.act_time))) from tblActivityTime as a1) " & _
			" And (select count(*) as theCount from tblActivityTime as b where (DATEPART (hh, b.act_time)) > (DATEPART (hh, getdate())))<=0 and act_Isactive = 1 order by act_time "
	adoRs.Open sSlotTrack,ConnFMG,3,3
	'Response.write sSlotTrack
	if adoRs.EOF = false then
		sTime = adoRs("act_Time")
		sDisplay = adoRs("act_Display")
		sDisp = FormatDateTime(sTime,3) & " - " & sDisplay
	if sTime <> "" then
		sActualTime = adoRs("act_Time")
	else
		sActualTime = ""
	end if

	else
		sActualTime = ""
	end if
%>
  <%
  	if sLocation <> "" then

	GetSlot = "select a.act_ID,b.slot_Time from tblActivityTime a, tblSlot_map b where a.act_ID = b.act_ID " & _
		" and a.act_Time = '"& sActualTime &"' and b.loc_ShortCode = '"&sLocation&"'"
	Set RsGetSlot = ConnFMG.Execute(GetSlot)
	'Response.write GetSlot
	if RsGetSlot.EOF = false then
		GSlot = RsGetSlot("slot_Time")
	end if
	RsGetSlot.Close

	strSlot = "select top 1 a.tdm_ID,a.act_Time from tblTransactionDetailsMaster a, tblSlot_Map b where a.loc_ShortCode = b.loc_ShortCode " & _
		" and b.loc_ShortCode = '"& sLocation &"' and b.Slot_Time = '"& GSlot &"' and datediff(d,a.tdm_date,'" & now() & "')= 0 and a.usr_EmpID = '"& Trim(session("sms_EmplNbr"))&"' " & _
		" order by a.tdm_ID desc"
	Set RsSlot = ConnFMG.Execute(strSlot)
	'Response.write strSlot
	if RsSlot.EOF = false then
		isExist = 1
		alert =  "The activity has already been filled for current slot."
		if RsSlot("act_Time") <> "" then
			sSlotTime = RsSlot("act_Time")
		else
			sSlotTime = ""
		end if

	end if
'	Response.write 	strSlot
'	Set objCheck  = Server.CreateObject("ADODB.Recordset")
'	sSlotCheck  = "Select act_Time from tblTransactionDetailsMaster where usr_EmpID = '"& Trim(Session("sms_EmplNbr"))&"' " & _
'		" and loc_ShortCode = '"& sLocation &"' and act_Time = '"& adoRs("act_Time")&"'  and datediff(d,tdm_date,'" & now() & "')=0  "
'	objCheck.Open sSlotCheck,ConnFmg,3,3
	'Response.write sSlotCheck & "<br>"
'	if objCheck.Eof = false then
'		isExist = 1
'		alert =  "The activity has already been filled for current slot."
'		if objCheck("act_Time") <> "" then
'			sSlotTime = objCheck("act_Time")
'		else
'			sSlotTime = ""
'		end if
'	end if
	'if sActualTime  = sSlotTime then
		'Response.write "Edit"
	'else
		'Response.write "ADD"
	'end if

	sGetEntryBy = "select tdm_EntryBy from tblTransactionDetailsMaster where loc_ShortCode = '"& sLocation &"' and act_Time = '"& sActualTime &"' " & _
		" and DateDiff(DD,tblTransactionDetailsMaster.tdm_Date,'"& now() &"') = 0 and usr_EmpID = '"& trim(session("sms_EmplNbr"))&"'"
	Set GetObj = ConnFMG.Execute(sGetEntryBy)
	if GetObj.EOF = false then
		sUser = GetObj("tdm_Entryby")
	else
		sUser = ""
	end if

'	objCheck.Close
'	Set objCheck = nothing

	end if
  %>

<html>
<head>
<title>Welcome to our Intranet | Health Check </title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
<Script language=JavaScript src="./includes/validate.js" type=text/javascript></SCRIPT>
<script language="javascript">
	function validateForm()
	{
	var frmObj = document.frmActivity;
	var x = frmObj.elements.length
	var count=0;
	if (frmObj.txt_Entryby.value == "")
	{
		alert("Please enter your name");
		frmObj.txt_Entryby.focus();
		return false;
	}

	for(i=0;i<x;i++)
	if(frmObj.elements[i].type=="radio")
	if (frmObj.elements[i].checked==true)
	{
		count++;
	}
	if (count==0)
	{
		alert("Please select checkpoint status");
		return false;
	}
	for(i=0;i<x;i++)
	if(frmObj.elements[i].type=="textarea")
	if (frmObj.elements[i].value.length > 1000)
	{
		alert("Comments should not exceed 1000 chars");
		frmObj.elements[i].focus()
		return false;
	}

	if (frmObj.txtSummary.value == "")
	{
		alert("Please enter Remarks");
		frmObj.txtSummary.focus()
		return false;
	}

//	alert("Done");
		frmObj.method="Post";
		frmObj.action="submit_helpdesk_activity.asp"
		frmObj.submit();
	}
	function openWinUplDoc()
	{
		window.open("user_upload_document.asp","UploadDocument","width=500,height=120,top=0,left=0,scrollbars=1");
	}
	function DataReload()
	{
		document.frmActivity.method="Post";
		document.frmActivity.action="helpdesk_activity.asp"
		document.frmActivity.submit();
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
              <tr>
                <td valign="top" height="1" colspan="5"><img src="images/blank.gif" border="0" width="1" height="1"></td>
              </tr>
              <tr>
                <td valign="top" width="5"><img src="images/blank.gif" border="0" width="15" height="1"></td>
                <td valign="top" width="80%" align="lefet"><img src="images/logo.jpg" border="0"></td>
                <td valign="bottom" align="center" width="5%">&nbsp;</td>
                <td valign="bottom" align="center" width="5%">&nbsp;</td>
                <td valign="bottom" align="center" width="5%">&nbsp;</td>
                <td valign="bottom" align="center" width="5%">&nbsp;</td>
                <td valign="top" width="5"><img src="images/blank.gif" border="0" width="1" height="1"></td>
              </tr>
              <tr>
                <td valign="top" height="4" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td>
              </tr>
              <tr>
                <td valign="top" bgcolor="#4FA3E4" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td>
              </tr>
              <tr>
                <td valign="top" height="2" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td>
              </tr>
              <tr>
                <td valign="top" colspan="6">
                  <table cellpadding="0" cellspacing="0" border="0" width="100%">
                    <tr>
                      <td valign="top"><img src="images/leftchk.gif" border="0"></td>
                      <td valign="top" width="1"><img src="images/blank.gif" border="0" width="1" height="1"></td>
                      <td valign="top" bgcolor="#3F89C3" width="70%" height="31"><img src="images/blank.gif" border="0" width="1" height="1"></td>
                      <td valign="top" width="1"><img src="images/blank.gif" border="0" width="1" height="1"></td>
                      <td valign="top"><img src="images/rightchk.gif" border="0"></td>
                      <td valign="middle" bgcolor="#3F89C3" width="20%" height="31" class="txtfnt" >&nbsp;&nbsp;<b>User:
                        <%=Session("sms_name") & " " & "(" & Session("sms_EmplNbr") & ")" %></b></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td valign="top" height="2" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td>
              </tr>
              <tr>
                <td valign="top" bgcolor="#4FA3E4" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td>
              </tr>
              <tr>
                <td valign="top" height="2" colspan="6"><img src="images/blank.gif" border="0" width="1" height="1"></td>
              </tr>
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
                        <form name="frmActivity" method="Post" action="submit_helpdesk_activity.asp" onSubmit="return validateForm();">
                          <table cellpadding="0" cellspacing="0" border="0" width="100%">
                            <tr>
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top" colspan="2">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3">
                                  <tr class="Content">
                                    <td width="40%" align="right" class="Content" bgcolor="#9FC4E1"><b>Location
                                      :</b> </td>
                                    <td bgcolor="#B4DAF8">
									<%
									Set objLoc = Server.CreateObject("ADODB.Recordset")
								'	sLoc = "select distinct a.loc_shortCode,b.loc_Name from tblTransaction as a inner join tblLocations b on  " & _
								'		" a.loc_ShortCode = b.loc_ShortCode where a.usr_EmpID = '" & Trim(Session("SMS_EmplNbr")) & "' "
									sLoc = "select distinct a.loc_shortCode,b.loc_Name from tblTransaction as a inner join tblLocations b on  " & _
											"  a.loc_ShortCode = b.loc_ShortCode inner join tblUsers c On a.loc_Shortcode = c.loc_Shortcode " & _
											"  where a.usr_EmpID = '" & Trim(Session("SMS_EmplNbr")) & "' and c.IsActive = 1 "
									objLoc.Open sLoc,ConnFMG,3,3
									%>
                                      <select name="cboLocation" onChange="DataReload(this.value)">
                                        <option value="0">Select Location</option>
									<%
										While Not ObjLoc.EOF
										if Trim(ObjLoc("loc_shortCode")) = Trim(sLocation) then
									%>
                                        <option value="<%=objLoc("loc_shortCode")%>" Selected><%=objLoc("loc_name")%></option>
									<%else%>
                                        <option value="<%=objLoc("loc_shortCode")%>"><%=objLoc("loc_name")%></option>
									<%end if%>
									<%
										objLoc.Movenext
										Wend
										objLoc.Close

									%>
                                      </select>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                              <td valign="top" width="5%">&nbsp;</td>
                            </tr>
                            <tr>
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top" colspan="2">&nbsp;</td>
                              <td valign="top" width="5%">&nbsp;</td>
                            </tr>
							<% if sLocation <> "0" and sLocation <> "" then %>
                            <tr>
                              <td valign="top" width="5%"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                              <td valign="top" colspan="2">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <tr align="center" class="Content" bgcolor="#FFFFFF">
                                    <td align="left" width="40%" class="txtHeader"><b>Activity</b>
                                    </td>
                                    <td align="right" class="txtHeader"><b>Date
                                      :</b> <%=adoRs("theDate")%></td>
                                  </tr>
                                  <tr align="center" class="Content">
                                    <td bgcolor="#9FC4E1" align="right" width="40%">Employee
                                      ID : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=objRs("usr_EmpID")%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#B4DAF8">
                                    <td align="right" bgcolor="#9FC4E1">Employee
                                      Name : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=objRs("usr_Name")%></td>
                                  </tr>
                                  <tr class="Content">
                                    <td align="right" bgcolor="#9FC4E1">Email
                                      : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=objRs("usr_Email")%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1">
                                    <td align="right">Location : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=objRs("loc_Name")%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1">
                                    <td align="right" bgcolor="#9FC4E1">Role :
                                    </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=objRs("rol_Name")%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1">
                                    <td align="right" bgcolor="#9FC4E1">Entry
                                      done by : </td>
                                    <td align="left" bgcolor="#B4DAF8">
                                      <input type="text" name="txt_Entryby" maxlength="100" value="<%=sUser%>">
                                      <font color="red">(Please enter your name)</font> </td>
                                  </tr>
                                  <tr class="Content" bgcolor="#FFFFFF" align="left">
                                    <td colspan="2"><font color="red"><%'=alert%></font></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#FFFFFF">
                                    <td align="right" colspan="2">
                                      <hr>
                                    </td>
                                  </tr>
								  <script language="javascript">
								  <% if sUser <> "" then %>
								  	alert ("you are editing activity details for" + " " + '<%=sDisp%>' + " " + "slot") ;
									<%else%>
									alert ("you are entering activity details for" + " " + '<%=sDisp%>' + " " + "slot") ;
									<%end if%>
								  </script>
                                  <tr class="Content">
                                    <td align="left" bgcolor="#FFFFFF" class="txtHeader"><b>Tracker</b></td>
                                    <td align="right" bgcolor="#FFFFFF" class="txtHeader"><b>Slot</b>
                                      : <%=FormatDateTime(adoRs("act_Time"),3) & " - " & adoRs("act_Display")%></td>
                                  </tr>
                                </table>
                              </td>
                              <td valign="top" width="5%"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                            </tr>
                            <tr>
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top" class="Content" colspan="2">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <%
									Set objLoop = Server.CreateObject("ADODB.RecordSet")
									sLoop = "Select a.usr_EmpID,a.area_ID,a.chktyp_ID,a.chkpnt_ID,a.tra_IsComments, b.area_Name,c.chktyp_Name,d.chkpnt_Name, a.loc_ShortCode " & _
										" from tblTransaction a Inner join tblAreas b On a.area_ID = b.area_ID Inner join tblChecktypes c ON a.chktyp_ID = c.chktyp_ID " & _
										" Inner join tblCheckpoint d On a.chkpnt_ID = d.chkpnt_ID where a.usr_EmpID = '"& Trim(session("sms_EmplNbr"))&"' and a.Loc_ShortCode = '"& Trim(sLocation)&"' " & _
										" order by a.area_ID,a.chktyp_ID,a.chkpnt_ID "
									objLoop.Open sLoop,ConnFMG,3,3

									area = ""
									checktype = ""
									checkpoint = ""
									While Not objLoop.EOF
								%>
                                  <% if area <> objLoop("area_name") then
								  	area = objLoop("area_name")
								  %>
                                  <tr align="center" class="Content" bgcolor="#9FC4E1">
                                    <td colspan="3"><b><%=area%></b></td>
                                  </tr>
                                  <%end if%>
                                  <%
								  	if checktype <> objLoop("chktyp_Name") then
									   checktype = objLoop("chktyp_Name")
								  %>
                                  <tr class="Content">
                                    <td colspan="3" align="left" bgcolor="#9FC4E1"><b><%=checktype%></b></td>
                                  </tr>
                                  <%end if%>
                                  <%
								  	if checkpoint <> objLoop("chkpnt_name") then
										checkpoint = objLoop("chkpnt_name")
								  %>
                                  <tr class="Content" bgcolor="#B4DAF8">
                                    <td align="left" bgcolor="#B4DAF8" width="40%" valign="top"><%=checkpoint%></td>
                                    <td valign="top" width="30%" >
                                      <%
										if objLoop("tra_IsComments") = "True" then
										sEdit =  "Select a.tdt_Comments,a.tdt_ActivityStatus,b.tdm_UploadDocument,b.tdm_Summary,b.tdm_Entryby " & _
											" from tblTransactionDetails a, tblTransactionDetailsMaster b where a.tdm_ID = b.tdm_ID and a.usr_EmpID = '"& Trim(session("sms_EmplNbr")) &"' " & _
											" and DateDiff(DD,a.tdt_Date,'"& Trim(now())&"')= 0 and a.Loc_ShortCode = '"&Trim(sLocation)&"' and a.area_ID = "& objLoop("area_ID") &" " & _
											" and a.chktyp_ID = "& objLoop("chktyp_ID") &" and  a.chkpnt_ID = "& objLoop("chkpnt_ID") &" and b.act_Time = '"& sActualTime &"' "
										Set RsEdit = ConnFMG.Execute (sEdit)
										'Response.write sEdit
										if RsEdit.Eof =  fasle then

											if RsEdit("tdt_Comments") <> "" then
												sComments	 = RsEdit("tdt_Comments")
											else
												sComments = ""
											end if
											if RsEdit("tdt_ActivityStatus") = "1" then
												sCheck = "1"
											elseif RsEdit("tdt_ActivityStatus") = "2" then
												sCheck = "2"
											else
												sCheck = "0"
											end if
											if RsEdit("tdm_UploadDocument") <> "" then
												sAttach = Trim(RsEdit("tdm_UploadDocument"))
											else
												sAttach = ""
											end if
										end if
									%>
                                      <textarea name="txtcomment_<%=objLoop("area_ID")%>_<%=objLoop("chktyp_ID")%>_<%=objLoop("chkpnt_ID")%>" rows="2" cols="25"><%=sComments%></textarea>
                                      <%
										end if
									%>
                                    </td>
                                    <td align="center">
                                      <input type="radio" name="rdo_<%=objLoop("area_ID")%>_<%=objLoop("chktyp_ID")%>_<%=objLoop("chkpnt_ID")%>" value="1" <% if sCheck = "1" then %> Checked <% end if %>>
                                      Success
                                      <input type="radio" name="rdo_<%=objLoop("area_ID")%>_<%=objLoop("chktyp_ID")%>_<%=objLoop("chkpnt_ID")%>" value="2" <% if sCheck = "2" then %> Checked <% end if %>>
                                      Failure </td>
                                  </tr>
                                  <%end if%>
                                  <%
								   sComments = ""
								   sCheck = ""
								  	objLoop.MoveNext
									Wend
								  %>
                                  <tr class="Content">
                                    <td bgcolor="#B4DAF8" align="right" width="40%">Attachment
                                      : </td>
                                    <td colspan="2" align="left" bgcolor="#B4DAF8">
                                      <input type="text" name="txtAttachment" maxlength="1000" size="30" readonly value="<%=sAttach%>">
                                      <input type="button" name="butDocument" value="Upload Document" onClick="openWinUplDoc();">
                                    </td>
                                  </tr>
								  <%
								  	sRemarks = "Select tdm_Summary from tblTransactionDetailsMaster where act_Time='"& sActualTime &"' and usr_EmpID = '"& Trim(session("sms_EmplNbr")) &"' " & _
										" and DateDiff(DD,tdm_Date,'"& Trim(now())&"')= 0 and Loc_ShortCode = '"&Trim(sLocation)&"'"
								  	Set RsRem = ConnFMG.Execute (sRemarks)

									if RsRem.EOF = false then
										sSummary = Rsrem("tdm_Summary")
									else
										sSummary = ""
									end if
								  %>
                                  <tr class="Content">
                                    <td bgcolor="#B4DAF8" align="right" width="40%" valign="top">Remarks: </td>
                                    <td colspan="2" align="left" bgcolor="#B4DAF8">
                                      <textarea name="txtSummary" cols="50" rows="3"><%=sSummary%></textarea>
                                    </td>
                                  </tr>

                                  <tr class="Content" bgcolor="#9FC4E1">
                                    <td colspan="3">&nbsp;</td>
                                  </tr>
                                  <% 'if isExist <> "1" then%>
                                  <tr class="Content" bgcolor="#9FC4E1">
                                    <td>
                                      <input type="hidden" name="hdSlot" value="<%=adoRs("act_Time")%>">
                                    </td>
                                    <td colspan="2">
                                      <input type="button" name="Submit" value="Submit" onClick="validateForm();">
                                      <input type="reset" name="Reset" value="Reset">
                                    </td>
                                  </tr>
                                  <% 'end if%>
                                </table>
                              </td>
                              <td valign="top" width="5%">&nbsp;</td>
                            </tr>
							<% end if %>
                            <tr>
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top" colspan="2">&nbsp;</td>
                              <td valign="top" width="5%">&nbsp;</td>
                            </tr>
                          </table>
                        </form>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
		</td>
		<!-- Middle Part End Here -->
		<td valign="top" bgcolor="#037DDB" width="1"><img src="./images/blank.gif" border="0" width="1" height="1"></td>

      <td valign="top" bgcolor="#F0F8FE" width="78"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
	</tr>
  </table>
</body>
</html>

<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="./includes/connect.asp"-->

<%
	Session.Timeout=40
Server.ScriptTimeout = 10000
	sLocation = Trim(Request.QueryString("slocation"))
	sSlot = Trim(Request.QueryString("sSlot"))
	sFilter = Trim(Request.QueryString("fdate"))
	'Response.write sLocation & ":" & sFilter & ":" & sSlot 
	
	EmpID = Trim(Session("sms_EmplNbr"))
		Set objRs =  Server.CreateObject("ADODB.Recordset")
		sQuery = "SELECT  tblRoles.rol_Name, tblLocations.loc_Name, tblLocations.loc_ShortCode, tblUsers.usr_EmpID, tblUsers.usr_Name, " & _ 
			" tblUsers.usr_Email,tblTransactionDetailsMaster.tdm_Entryby  FROM  tblUsers INNER JOIN tblRoles ON tblUsers.rol_ID = tblRoles.rol_ID " & _ 
			" INNER JOIN tblLocations ON tblUsers.loc_ShortCode = tblLocations.loc_ShortCode Inner join tblTransactionDetailsMaster ON " & _ 
			" tblTransactionDetailsMaster.loc_ShortCode = tblLocations.loc_ShortCode WHERE (tblUsers.loc_ShortCode = '"& sLocation &"' " & _ 
			" and tblTransactionDetailsMaster.act_time = '"& sSlot &"' and DateDiff(DD,tblTransactionDetailsMaster.tdm_Date,'"&sFilter&"') = 0) " & _ 
			" and tblRoles.rol_ID = 3 "
		objRs.Open sQuery,ConnFMG,3,3
'		Response.write 	sQuery
		if objRs.Eof = false then
			if objRs("tdm_Entryby") <> "" then
				EntryBy = objRs("tdm_Entryby")
			else
				EntryBy = ""
			end if
			if objRs("usr_EmpID") <> "" then
				usr_EmpID = objRs("usr_EmpID")
			else
				usr_EmpID = ""
			end if
			if objRs("usr_Name") <> "" then
				usr_Name = objRs("usr_Name")
			else
				usr_Name = ""
			end if
			if objRs("usr_Email") <> "" then
				usr_Email = objRs("usr_Email")
			else
				usr_Email = ""
			end if
			if objRs("loc_Name") <> "" then
				loc_Name = objRs("loc_Name")
			else
				loc_Name = ""
			end if
			if objRs("rol_Name") <> "" then
				rol_Name = objRs("rol_Name")
			else
				rol_Name = ""
			end if
			
		end if	
%>
 <%
	Set adoRs = Server.CreateObject("ADODB.Recordset")
	sSlotTrack = " Select top 1 *,GetDate() as theDate from tblActivityTime where (DATEPART (hh, act_time)) > (DATEPART (hh, getdate())) and act_Isactive = 1 " & _ 
			" Union ALL " & _ 
			" Select top 1 a.*,DateAdd(DD,1,GetDate()) as theDate from tblActivityTime as a where (DATEPART (hh, a.act_time)) = (Select MIN((DATEPART (hh, a1.act_time))) from tblActivityTime as a1) " & _ 
			" And (select count(*) as theCount from tblActivityTime as b where (DATEPART (hh, b.act_time)) > (DATEPART (hh, getdate())))<=0 and act_Isactive = 1 order by act_time "
	adoRs.Open sSlotTrack,ConnFMG,3,3
	
%>

<html>
<head>
<title>Welcome to our Intranet | Health Check of Datacenter</title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
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
                      <td valign="middle" bgcolor="#3F89C3" width="20%" height="31" class="txtfnt" >&nbsp;&nbsp;<b>User
                        : <%=Session("sms_name") & " " & "(" & Session("sms_EmplNbr") & ")" %></b></td>
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

                    <tr>
                      <td valign="top" colspan="3"><br>
                      </td>
                    </tr>
                    <tr>
                      <td valign="top"  colspan="3">
                          <table cellpadding="0" cellspacing="0" border="0" width="100%">
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
                                    <td align="left" width="40%" class="txtHeader"><b>Employee 
                                      Details </b></td>
                                    <td align="right" class="txtHeader"><b>Date 
                                      :</b> <%=FormatDateTime(sFilter,1)%></td>
                                  </tr>
                                  <tr align="center" class="Content"> 
                                    <td bgcolor="#9FC4E1" align="right" width="40%">Employee 
                                      ID : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=usr_EmpID%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#B4DAF8"> 
                                    <td align="right" bgcolor="#9FC4E1">Employee 
                                      Name : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=usr_Name%></td>
                                  </tr>
                                  <tr class="Content"> 
                                    <td align="right" bgcolor="#9FC4E1">Email 
                                      : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=usr_Email%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1"> 
                                    <td align="right">Location : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=loc_Name%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1"> 
                                    <td align="right" bgcolor="#9FC4E1">Role : 
                                    </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=rol_Name%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1"> 
                                    <td align="right" bgcolor="#9FC4E1">Entry 
                                      done by : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=EntryBy%>
                                  </tr>
                                  <tr class="Content" bgcolor="#FFFFFF"> 
                                    <td align="right" colspan="2"> 
                                      <hr>
                                    </td>
                                  </tr>
                                  <tr class="Content"> 
                                    <td align="left" bgcolor="#FFFFFF" class="txtHeader"><b>Activity 
                                      Details </b></td>
                                    <td align="right" bgcolor="#FFFFFF" class="txtHeader"><b>Slot</b> 
                                      : <%=FormatDateTime(sSlot,3)%></td>
                                  </tr>
                                </table>
                              </td>
                              <td valign="top" width="5%"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                            </tr>
                            <tr> 
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top" class="Content" colspan="2"> 
								<%
									Set objLoop = Server.CreateObject("ADODB.RecordSet")
									sLoop = "Select a.usr_EmpID,a.area_ID,a.chktyp_ID,a.chkpnt_ID,a.tra_IsComments, b.area_Name,c.chktyp_Name,d.chkpnt_Name, a.loc_ShortCode " & _ 
										" from tblTransaction a Inner join tblAreas b On a.area_ID = b.area_ID Inner join tblChecktypes c ON a.chktyp_ID = c.chktyp_ID " & _ 
										" Inner join tblCheckpoint d On a.chkpnt_ID = d.chkpnt_ID where a.usr_EmpID = '"& Trim(usr_EmpID)&"' and a.Loc_ShortCode = '"& Trim(sLocation)&"' " & _ 
										" order by a.area_ID,a.chktyp_ID,a.chkpnt_ID "
								'	sLoop = "Select a.usr_EmpID,a.area_ID,a.chktyp_ID,a.chkpnt_ID,a.tra_IsComments, b.area_Name,c.chktyp_Name,d.chkpnt_Name, " & _ 
								'		" a.loc_ShortCode,e.tdt_Comments,e.tdt_ActivityStatus from tblTransaction a Inner join tblAreas b On a.area_ID = b.area_ID " & _ 
								'		" Inner join tblChecktypes c ON a.chktyp_ID = c.chktyp_ID Inner join tblCheckpoint d On a.chkpnt_ID = d.chkpnt_ID Inner join " & _ 
								'		" tblTransactionDetails e ON a.chkpnt_ID = e.chkpnt_ID where a.usr_EmpID = '"& Trim(usr_EmpID) &"' and DateDiff(DD,e.tdt_Date,'"& Trim(sFilter)&"')=0  " & _ 
								'		" and a.Loc_ShortCode = '"&Trim(sLocation)&"' order by a.area_ID,a.chktyp_ID,a.chkpnt_ID "
										
									objLoop.Open sLoop,ConnFMG,3,3
								'	Response.write sLoop
								%>
								<% if objLoop.EOF = false then %>
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <%
									area = ""
									checktype = ""
									checkpoint = ""
									While Not objLoop.EOF											
								%>
                                  <%if area <> objLoop("area_name") then
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
                                    
                                  <td valign="top" width="40%" bgcolor="#B4DAF8" > 
                                    <%
										if objLoop("tra_IsComments") = "True" then
										sEdit =  "Select a.tdt_Comments,a.tdt_ActivityStatus,b.tdm_UploadDocument,b.tdm_Summary,b.tdm_Entryby " & _ 
											" from tblTransactionDetails a, tblTransactionDetailsMaster b where a.tdm_ID = b.tdm_ID and a.usr_EmpID = '"& Trim(usr_EmpID) &"' " & _ 
											" and DateDiff(DD,a.tdt_Date,'"& Trim(sFilter)&"')= 0 and a.Loc_ShortCode = '"&Trim(sLocation)&"' and a.area_ID = "& objLoop("area_ID") &" " & _
											" and a.chktyp_ID = "& objLoop("chktyp_ID") &" and  a.chkpnt_ID = "& objLoop("chkpnt_ID") &" and b.act_Time = '"& sSlot &"' " 

										Set RsEdit = ConnFMG.Execute (sEdit)
										
										if RsEdit.Eof =  fasle then
										
											if RsEdit("tdt_Comments") <> "" then
												sComments	 = RsEdit("tdt_Comments")
											else
												sComments = ""	
											end if
											
											if RsEdit("tdt_ActivityStatus") = "1" then
												sCheck = "1"
												color = "#339933"
											elseif RsEdit("tdt_ActivityStatus") = "2" then
												sCheck = "2"
												color = "#FF0000"
											else
												sCheck = "0"
												color="#4682b4"
											end if
											if RsEdit("tdm_UploadDocument") <> "" then
											
												sAttach = "<a href='"&Trim(RsEdit("tdm_UploadDocument"))&"' target=_blank class=Content>View Attachment</a> "
											else
												sAttach = "No attachment found"
											end if
											
											
										end if
									%>
                                    <textarea name="txtcomment_<%=objLoop("area_ID")%>_<%=objLoop("chktyp_ID")%>_<%=objLoop("chkpnt_ID")%>" rows="2" cols="40" readonly><%=sComments%></textarea>
                                    <%
										end if
									%>
                                  </td>
                                    
                                  <td align="center" width="5%" bgcolor="<%=color%>">&nbsp;								  
								  </td>
                                  </tr>
                                  <%end if%>
                                  <%
								  	color="#4682b4"
								  	sCheck = ""
									sComments = ""
								  	objLoop.MoveNext
									Wend
								  %>
                                  <tr class="Content"> 
                                    <td bgcolor="#B4DAF8" align="right" width="40%">Attachment 
                                      : </td>
                                  <td colspan="2" align="left" bgcolor="#B4DAF8"><%=sAttach%>
                                  </td>
                                  </tr>
								  <%
								  	sRemarks = "Select tdm_Summary from tblTransactionDetailsMaster where act_Time='"& sSlot&"' and usr_EmpID = '"& Trim(usr_EmpID) &"' " & _ 
										" and DateDiff(DD,tdm_Date,'"& Trim(sFilter)&"')= 0 and Loc_ShortCode = '"&Trim(sLocation)&"'"
								  	Set RsRem = ConnFMG.Execute (sRemarks)	
									if RsRem.EOF = false then
										Summary = Rsrem("tdm_Summary")
									end if
								  %>
                                  <tr class="Content"> 
                                    <td bgcolor="#B4DAF8" align="right" width="40%" valign="top">Remarks 
                                      : </td>
                                    <td colspan="2" align="left" bgcolor="#B4DAF8"> 
                                      <textarea name="txtSummary" cols="50" rows="3" readonly><%=Summary%></textarea>
                                    </td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1"> 
                                    <td colspan="3">&nbsp;</td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1"> 
                                    <td colspan="3">&nbsp; </td>
                                  </tr>
                                </table>
								<% else %>
								<table width="100%"><tr bgcolor="#9FC4E1"><td align="center" class="Content"><b>No Details Exist</b></td></tr></table>
								<%end if%>
                              </td>
                              <td valign="top" width="5%">&nbsp;</td>
                            </tr>
							<%end if%>
                            <tr> 
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top" colspan="2">&nbsp;</td>
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
      </table>
		</td>
		<!-- Middle Part End Here -->
		<td valign="top" bgcolor="#037DDB" width="1"><img src="./images/blank.gif" border="0" width="1" height="1"></td>

      <td valign="top" bgcolor="#F0F8FE" width="78"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
	</tr>
  </table>
</body>
</html>

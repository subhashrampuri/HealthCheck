<%@LANGUAGE="VBSCRIPT"%>
<% if  not (session("sms_validated")) = "True" then
	session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
	Response.Redirect "/default.asp"
 	end if
	Session.Timeout=40
	lclint_currPage = request("pno")
	if lclint_currPage = "" then
		lclint_currPage = 1
	end if

%> 

<!--#include file="./includes/connect.asp"-->

<%
lclint_MesgNo = Request.QueryString("msgno")
lclstr_ActionType = REPLACE(Request.QueryString("actionType"),"'","")
Chkpointid = Request.Querystring("chkpntid")

if UCASE(lclstr_ActionType) = UCASE("UPDATE") then

	If Request.Querystring("chkpntid") <> "" Then
		Chkpointid = Request.Querystring("chkpntid")
		lclstr_SQL = "select chkpnt_ID,chkpnt_name,area_ID,chktyp_ID,isActive from tblCheckpoint where chkpnt_ID=" & Chkpointid & ""
		Set lclobj_altuserRec = Server.Createobject("ADODB.Recordset")
		lclobj_altuserRec.Open lclstr_SQL,ConnFMG,3,3

		If Not (lclobj_altuserRec.EOF and lclobj_altuserRec.BOF) Then
			lclint_areaid = lclobj_altuserRec("area_ID")
			lclint_chktypeid = lclobj_altuserRec("chktyp_ID")
			'lclint_altuser = lclobj_altuserRec("alt_EmpID")
			lclstr_chkpntname = lclobj_altuserRec("chkpnt_name")
			if lclobj_altuserRec("isActive") = "True" then
				lclstr_isactive = "Checked"
			Else
				lclstr_isactive = ""
			end if
		End If 
		lclobj_altuserRec.close()
		Set lclobj_altuserRec = Nothing
		
		
		sMap = "Select alt_EmpID from tblCheckpoint_Map_Alertuser where chkpnt_ID = "& Chkpointid &""
		Set MapRs = Server.Createobject("ADODB.Recordset")
		MapRs.Open sMap,ConnFMG,3,3
		vArray = ""				
		while NOT MapRs.EOF 
			vArray =   vArray & Trim(MapRs("alt_EmpID")) & ","
			'Response.write vArray
		MapRs.MoveNext
		 'vArray &
		 
		Wend
			
		
	End If
elseif UCASE(lclstr_ActionType) = UCASE("ADD") then
	lclstr_HRName = REPLACE(Request.QueryString("area_name"),"'","")
end if
if lclstr_ActionType = "" then
lclstr_ActionType = "ADD"
end if
%>
<html>
<head>
<title>Welcome to our Intranet | Health Check of Datacenter</title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
<Script language=JavaScript src="./includes/validate.js" type=text/javascript></SCRIPT>
<script language="JavaScript">
<!--
	function validateForm()
	{
		var frmObj = document.frmCheckpoint ;
		var mesg = "" ;
		var flag = 0 ;

		if(frmObj.cboArea.value == "0")
		{
			flag = 1 ;
			mesg = mesg + "Please select area\n" ;
			frmObj.cboArea.focus();
		}
		if(frmObj.cboChktype.value == "0")
		{
			flag = 1 ;
			mesg = mesg + "Please select checktype \n" ;
			frmObj.cboChktype.focus();
		}
		if(frmObj.txtChkpoint.value == "")
		{
			flag = 1 ;
			mesg = mesg + "Check point is required field \n" ;
			frmObj.txtChkpoint.focus();
		}
		else if (isWhitespace(frmObj.txtChkpoint.value))
		{
			flag = 1 ;
			mesg = mesg + "Check point is required field \n" ;
			frmObj.txtChkpoint.focus();
		}
	/*	if(frmObj.txtaltusers.value == "0")
		{
			flag = 1 ;
			mesg = mesg + "Please select User\n" ;
			frmObj.txtaltusers.focus();
		}
	*/	if(flag == 1)
		{
			alert(mesg);
			return false ;
		}
		return true ;
	}
//-->
</script>		

</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#ffffff">
  <table cellpadding="0" cellspacing="0" border="0" height="100%" bgcolor="#ffffff" >
    <tr>
		
      
    <td valign="top" bgcolor="#F0F8FE" width="59" ><img src="./images/blank.gif" border="0" width="1" height="1"></td>
    <td valign="top" bgcolor="#037DDB" width="1"  ><img src="./images/blank.gif" border="0" width="1" height="1"></td>
		<!-- Middle Part Start Here -->
		
      
    <td valign="top" bgcolor="#FFFFFF" width="879" > 
      <table cellpadding="0" cellspacing="0" border="0" width="100%"  vAlign="top">
               <tr>
			   		<td valign="top">
						<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff"  vAlign="top">
						    <tr><td valign="top" height="1" colspan="5"><img src="images/blank.gif" border="0" width="1" height="1"></td></tr>
							<tr>
								<td valign="top" width="5"><img src="images/blank.gif" border="0" width="15" height="1"></td>
								<td valign="top" width="80%" align="left"><img src="images/logo.jpg" border="0"></td>
								
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
											<td valign="top" width="1"><img src="images/blank.gif" border="0" width="1" height="1"></td>
											<td valign="top"><img src="images/rightchk.gif" border="0"></td>
											<td valign="middle" bgcolor="#3F89C3" width="20%" height="31" class="txtfnt">&nbsp;&nbsp;<b>User: <%=Session("sms_name") & " " & "(" & Session("sms_EmplNbr") & ")" %></b></td>	
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
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                            <td valign="top" class="txtHeader"><b>&nbsp;Add Check 
                              Points</b><br>
                              <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                            </td>
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                          </tr>
                          <tr> 
                            <td valign="top"> 
                              <form name="frmCheckpoint" method="post" action="checkpoint_submit.asp" onSubmit="return validateForm();">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1" width="40%" valign="middle"> 
                                      <div align="right">Area :</div>
                                    </td>
                                    <td bgcolor="#B4DAF8" valign="top" align="left"> 
                                      <%
										Set objRs=Server.Createobject("ADODB.Recordset")
										MySql= "Select area_ID,area_Name from tblAreas where isActive =1 Order by area_name"
										objRs.Open MySql,ConnFMG,3,3
									 %>
                                      <div align="left"> 
                                        <select name="cboArea" id="cboArea" style="width:220px">
                                          <option value="0">Select Area</option>
                                          <%
												While Not objRs.EOF
												if lclint_areaid  = cInt(objRs("area_ID")) then
													str = "selected"
												else
													str = ""	
												end if
													
											%>
                                          <option value="<%=objRs("area_ID")%>" <%=str%>><%=objRs("area_Name")%></option>
                                          <%
												objRS.MoveNext
												Wend
											%>
                                        </select>
                                      </div>
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1" valign="middle"> 
                                      <div align="right">Check Type : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8" valign="top" align="left"> 
                                      <%
										Set adoRs=Server.Createobject("ADODB.Recordset")
										MySql= "Select chktyp_ID,chktyp_Name from tblChecktypes where isActive = 1 Order by chktyp_Name"
										adoRs.Open MySql,ConnFMG,3,3
									 %>
                                      <div align="left"> 
                                        <select name="cboChktype" id="cboChktype" style="width:220px">
                                          <option value="0">Select Checktype</option>
                                          <%
												While Not adoRs.EOF
												if lclint_chktypeid = adoRs("chktyp_ID") then
													mystr = "selected"
												else
													mystr = ""	
												end if
											%>
                                          <option value="<%=adoRs("chktyp_ID")%>" <%=mystr%>><%=adoRs("chktyp_Name")%></option>
                                          <%
												adoRs.MoveNext
												Wend
												set adoRs = Nothing
											%>
                                        </select>
                                      </div>
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1" valign="middle"> 
                                      <div align="right">Check Points : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8" valign="top" align="left"> 
                                      <div align="left"> 
                                        <input type="text" name="txtChkpoint" size="29" id="txtChkpoint" maxlength=1500 style="width:220px" value="<%=lclstr_chkpntname%>">
                                      </div>
                                    </td>
                                  </tr>
                                  <!--<tr class="Content" > 
                                    <td bgcolor="#9FC4E1" valign="top"> 
                                      <div align="right">Alert Users : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8" valign="top" align="left"> 
                                      <%
										Set lclobj_empRec=Server.Createobject("ADODB.Recordset")
										lclstr_SQL = "select alt_EmpID,alt_Name from tblAlertUsers where isActive=1" 
										lclobj_empRec.Open lclstr_SQL,ConnFMG,3,3
									%>
                                      <div align="left"> 
                                        <select name="txtaltusers" id="txtaltusers" size="4" style="width:220px" multiple>
                                          <%
										While Not lclobj_empRec.EOF
											Arry = Split(vArray,",")
											For i = 0 to UBound(Arry) - 1
											
												if cStr(Arry(i)) = cStr(lclobj_empRec("alt_EmpID")) then
													lclstr_print = "Selected"
													exit for
												else
													lclstr_print = ""	
												end if
												
											Next
										
									'	if lclint_altuser = lclobj_empRec("alt_EmpID") then
									'		lclstr_print = "selected"
									'	else
									'		lclstr_print = ""	
									'	end if
								%>
                                          <option value="<%=lclobj_empRec("alt_EmpID")%>" <%=lclstr_print%>><%=lclobj_empRec("alt_Name")%></option>
                                          <%
										lclobj_empRec.MoveNext
										Wend
										set lclobj_empRec = Nothing
								%>
                                        </select>
                                      </div>
                                    </td>
                                  </tr>-->
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1" valign="middle"> 
                                      <div align="right">Is Active : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8" valign="top" align="left"> 
                                      <div align="left"> 
                                        <input type="checkbox" name="IsActive" value="1" <%=lclstr_isactive%>>
                                      </div>
                                    </td>
                                  </tr>
                                  <tr class="Content"> 
                                    <td bgcolor="#9FC4E1">&nbsp;</td>
                                    <td bgcolor="#B4DAF8"> 
                                      <input type="submit" name="butSubmit" value="<%=UCASE(lclstr_ActionType)%>">
                                      <input type="reset" name="Reset" value="Reset">
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td> 
                                      <input type="hidden" name="hdchkpnt_ID" id="hdchkpnt_ID" value="<%=Chkpointid%>">
                                    </td>
                                    <td> <span  align=center class="Content"><font color="red" size=2> 
                                      <%
										if lclint_MesgNo = 1 then
											Response.Write "Checkpoint Added Successfully"
										elseif lclint_MesgNo = 2 then
											Response.Write "Sorry! Cannot Add Duplicate Checkpoint"
										elseif lclint_MesgNo = 3 then
											Response.Write "Checkpoint Updated Successfully"
										end if
										%>
                                      </font></span> </td>
                                  </tr>
                                </table>
                              </form>
                            </td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top"> 
                              <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                <tr align="left" class="Content"> 
                                  <td bgcolor="#FFFFFF" colspan="7" class="txtHeader"><b> 
                                    &nbsp;Manage Check Points</b></td>
                                </tr>
                                <tr align="center" class="Content"> 
                                  <td bgcolor="#9FC4E1" width="5%"><b>Sl. No</b></td>
                                  <td bgcolor="#9FC4E1"><b>Check Points</b></td>
                                  <td bgcolor="#9FC4E1"><b>Check Types</b></td>
                                  <td bgcolor="#9FC4E1"><b>Area</b></td>
                                  <!--<td bgcolor="#9FC4E1"><b>Alert User</b></td>-->
                                  <td bgcolor="#9FC4E1"><b>Is Active</b></td>
                                  <td bgcolor="#9FC4E1"><b>Edit</b></td>
                                </tr>
                                <%
									Set adoRs = Server.CreateObject("ADODB.RecordSet")
									MySql = "select a.chkpnt_ID,a.chkpnt_name,a.area_ID,b.area_name,c.chktyp_ID,c.chktyp_Name,a.isActive from tblCheckpoint " & _ 
										" a join tblAreas b on b.area_id=a.area_id  join tblChecktypes c on c.chktyp_id=a.chktyp_id " & _
										" Order by a.chkpnt_name,c.chktyp_Name,b.area_name"

									adoRs.Open MySql,ConnFMG,3,3
									if adoRs.Eof = false then
									i  = 1
									color="#9FC4E1"

									lclint_MaxPageSize = 10
									adoRs.PageSize = lclint_MaxPageSize
									lclint_MaxPageCount = adoRs.PageCount
									adoRs.AbsolutePage = lclint_currPage

									lclint_listNum = 0
									i = 1
									While Not (adoRs.Eof or (lclint_listNum+1) > adoRs.PageSize)
									lclint_listNum = lclint_listNum + 1
									
									if color="#9FC4E1" then
										color="#B4DAF8"
									else
										color="#9FC4E1"
									end if	
									
									if adors("isActive") = "True" then
										sActive = "Active"
									else
										sActive = "In Active"									
									end if
									chkpnt = adors("chkpnt_ID")
								%>
                                <tr class="Content" bgcolor=<%=color%>> 
                                  <td align="center"><%=i%></td>
                                  <td align="left" style="word-break: break-all; width:300px;"><%=adoRs("chkpnt_name")%></td>
                                  <td align="left" style="word-break: break-all; width:300px;"><%=adoRs("chktyp_Name")%></td>
                                  <td align="left" style="word-break: break-all; width:180px;"><%=adoRs("area_name")%></td>
                                  <%
								  	sCheck = "Select chkpnt_ID from tblCheckpoint_Map_Alertuser Where chkpnt_ID = "& chkpnt &" "
									Set ChkRs = Server.CreateObject("ADODB.Recordset")
									ChkRs.Open sCheck,ConnFMG,3,3
									if ChkRs.EOF = False then
										sMark = "Yes"
									else
										sMark = "No"
									end if
									Set ChkRs = Nothing
								  %>
                                  <!--<td align="center"><%=sMark%></td>-->
                                  <td align="center"><%=sActive%></td>
                                  <td align="center"><a class="content" href="manage_admin_checkpoints.asp?chkpntid=<%=adoRs("chkpnt_ID")%>&actionType=Update">Edit</a></td>
                                </tr>
                                <% 
									i = i + 1
									adoRs.MoveNext
									wend
								else%>
                                <tr class="Content" bgcolor="#B4DAF8" align="center"> 
                                  <td colspan="7" bgcolor="#FFFFFF">No Records 
                                    Found</td>
                                </tr>
                                <%	
								end if
									
								%>
                                <tr class="Content" bgcolor="#B4DAF8" align="right"> 
                                  <td colspan="7" bgcolor="#FFFFFF"> 
                                    <div align="right">Page : 
                                      <%
										lclint_pageNum = 1
										while lclint_pageNum <= lclint_MaxPageCount
											if CINT(lclint_currPage) = lclint_pageNum then
											Response.Write lclint_currPage
											else
											%>
                                      <A  href="manage_admin_checkpoints.asp?pno=<%=lclint_pageNum%>" ><%=lclint_pageNum%></a> 
                                      <%
											end if
											Response.Write "&nbsp;"
										lclint_pageNum = lclint_pageNum+1
										wend
									%>
                                    </div>
                                  </td>
                                </tr>
                              </table>
                            </td>
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
		
    <td valign="top" bgcolor="#037DDB" width="1" ><img src="./images/blank.gif" border="0" width="1" height="1"></td>
		
      
    <td valign="top" bgcolor="#F0F8FE" width="78" ><img src="./images/blank.gif" border="0" width="1" height="1"></td>
	</tr>
  </table>

</body>
</html>

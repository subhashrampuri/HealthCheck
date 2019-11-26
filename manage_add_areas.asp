<%@LANGUAGE="VBSCRIPT"%>
<% if  not (session("sms_validated")) = "True" then
	session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
	Response.Redirect "/default.asp"
 	end if
		Session.Timeout=40
%> 
<!--#include file="./includes/connect.asp"-->
<%
'Response.write sCheckLogin
'lclstr_chkIsActive = true
lclstr_disable = ""
lclstr_ActionType = REPLACE(Request.QueryString("actionType"),"'","")
lclint_MesgNo = Request.QueryString("msgno")
lclstr_areaID = REPLACE(Request.QueryString("area_ID"),"'","")

if UCASE(lclstr_ActionType) = UCASE("UPDATE") then
	lclstr_disable = "READONLY"
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	lclstr_mysql ="select area_Name,isActive from tblAreas where area_ID=" & lclstr_areaID & " "
	'response.write lclstr_mysql
	objRs.Open lclstr_mysql,ConnFMG,3,3
	if not objRs.eof then
		AreaName = objRs("area_Name")
		if objRs("isActive") = "True" then
			Active = "Checked"
		end if
		
	End if
	Set objRs = NOTHING
elseif UCASE(lclstr_ActionType) = UCASE("ADD") then
	lclstr_HRName = REPLACE(Request.QueryString("area_name"),"'","")
end if
if lclstr_ActionType = "" then
lclstr_ActionType = "ADD"
end if
	
%>
<html>
<head>
<title>Welcome to our Intranet | Health Check</title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
<Script language=JavaScript src="./includes/validate.js" type=text/javascript></SCRIPT>
<script language="JavaScript">
<!--
	function validateForm()
	{
		var frmObj = document.frmArea ;
		var mesg = "" ;
		var flag = 0 ;

		if(frmObj.txtArea.value == "")
		{
			flag = 1 ;
			mesg = mesg + "Area name is required field\n" ;
			frmObj.txtArea.focus();
		}
		else if (isWhitespace(frmObj.txtArea.value))
		{
			flag = 1 ;
			mesg = mesg + "Area name is required field\n" ;
			frmObj.txtArea.focus();
		}
		if(flag == 1)
		{
			alert(mesg);
			return false ;
		}
		return true ;
	}
//-->
</script>
		
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#ffffff" >
  <table cellpadding="0" cellspacing="0" border="0" height="100%" bgcolor="#ffffff">
    <tr>

      <td valign="top" bgcolor="#F0F8FE" width="59"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
		<td valign="top" bgcolor="#037DDB" width="1" ><img src="./images/blank.gif" border="0" width="1" height="1"></td>

      <td valign="top" bgcolor="#FFFFFF" width="879">
        <table cellpadding="0" cellspacing="0" border="0" width="100%"  vAlign="top">
               <tr>
			   		<td valign="top">
						<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" vAlign="top">
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
											<td valign="top" width="1"><img src="images/blank.gif" border="0" width="1" height="1"></td>
											<td valign="top"><img src="images/rightchk.gif" border="0"></td>

                        <td valign="middle" bgcolor="#3F89C3" width="20%" height="31" class="txtfnt" >&nbsp;&nbsp;<b>User:
                          <%=Session("sms_name") & " " & "(" & Session("sms_EmplNbr") & ")" %> </b></td>
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
                            <td valign="top" class="txtHeader"><b>&nbsp;Add Areas</b><br>
                              <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                            </td>
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                          </tr>
                          <tr> 
                            <td valign="top"> 
                              <form name="frmArea" method="post" action="area_submit.asp" onSubmit="return validateForm();">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1" width="40%"> 
                                      <div align="right">Areas : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8"> 
                                      <div align="left"> 
                                        <input type="text" name="txtArea" maxlength="100" value="<%=AreaName%>">
                                      </div>
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1"> 
                                      <div align="right">Is Active : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8"> 
                                      <div align="left"> 
                                        <input type="checkbox" name="IsActive" value="1" <%=Active%>>
                                      </div>
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1">&nbsp;</td>
                                    <td bgcolor="#B4DAF8"> 
                                      <input type="submit" name="butSubmit" value="<%=UCASE(lclstr_ActionType)%>">
                                      <input type="reset" name="Reset" value="Reset">
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td> 
                                      <input type="hidden" name="hdarea_ID" id="hdarea_ID" value="<%=lclstr_areaID%>">
                                    </td>
                                    <td> <span  align=center class="Content"><font color="red" size=2> 
                                      <%
										if lclint_MesgNo = 1 then
											Response.Write "Area Added Successfully"
										elseif lclint_MesgNo = 2 then
											Response.Write "Sorry! Cannot Add Duplicate Area"
										elseif lclint_MesgNo = 3 then
											Response.Write "Area Updated Successfully"
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
                              <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                <tr align="left" class="Content" bgcolor="#FFFFFF"> 
                                  <td colspan="4" class="txtHeader"><b>Manage 
                                    Areas</b></td>
                                </tr>
                                <tr align="center" class="Content"> 
                                  <td bgcolor="#9FC4E1" width="15%"><b>Sl. No</b></td>
                                  <td bgcolor="#9FC4E1"><b>Area Name</b></td>
                                  <td bgcolor="#9FC4E1"><b>Status</b></td>
                                  <td bgcolor="#9FC4E1"><b>Edit</b></td>
                                </tr>
                                <%
									Set adoRs = Server.CreateObject("ADODB.RecordSet")
									MySql = " select area_ID,area_Name,isActive from tblAreas order by area_Name"
									adoRs.Open MySql,ConnFMG,3,3
									if adoRs.Eof = false then
									color="#9FC4E1"
									i = 1
										While Not adoRs.Eof
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
										
								%>
                                <tr class="Content" bgcolor=<%=color%>> 
                                  <td align="center"><%=i%></td>
                                  <td align="left" ><%=adoRs("area_Name")%> </td>
                                  <td align="center"><%=sActive%></td>
                                  <td align="center"><a class="content" href="manage_add_areas.asp?area_ID=<%=adors("area_ID")%>&actionType=Update">Edit</a></td>
                                </tr>
                                <% 
									i = i + 1
									adoRs.MoveNext
									wend
								else %>
								
                                <tr class="Content" bgcolor="#B4DAF8" align="center"> 
                                  <td colspan="5" bgcolor="#FFFFFF">No Records 
                                    Found</td>
                                </tr>
								<%	
								end if
									
								%>
                                <tr class="Content" bgcolor="#B4DAF8" align="right">
                                  <td colspan="5" bgcolor="#FFFFFF">&nbsp;</td>
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
		<td valign="top" bgcolor="#037DDB" width="1"><img src="./images/blank.gif" border="0" width="1" height="1"></td>

      <td valign="top" bgcolor="#F0F8FE" width="78"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
	</tr>
  </table>

</body>
</html>

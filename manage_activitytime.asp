<%@LANGUAGE="VBSCRIPT"%>
<% if  not (session("sms_validated")) = "True" then
	session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
	Response.Redirect "/default.asp"
 	end if
	Session.Timeout=40

%>
<!--#include file="./includes/connect.asp"-->
<%
lclstr_ActionType = REPLACE(Request.QueryString("actionType"),"'","")
lclint_MesgNo = Request.QueryString("msgno")
lclstr_actID = REPLACE(Request.QueryString("act_ID"),"'","")

if UCASE(lclstr_ActionType) = UCASE("UPDATE") then
	lclstr_disable = "READONLY"
	Set adoRs = Server.CreateObject("ADODB.RecordSet")
	lclstr_mysql ="select act_ID,act_Time,act_isActive from tblActivityTime where act_ID=" & lclstr_actID & " "
	'response.write lclstr_mysql
	adoRs.Open lclstr_mysql,ConnFMG,3,3
	if not adoRs.eof then
		act_Time = adoRs("act_Time")

		if adoRs("act_isActive") = "True" then
			Active = "Checked"
		end if

	End if
	Set adoRs = NOTHING
elseif UCASE(lclstr_ActionType) = UCASE("ADD") then
	lclstr_HRName = REPLACE(Request.QueryString("act_Time"),"'","")
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
		var frmObj = document.frmChecktype ;
		var mesg = "" ;
		var flag = 0 ;

		if(frmObj.cboTime.value == "100")
		{
			flag = 1 ;
			mesg = mesg + "Please select time\n" ;
			frmObj.cboTime.focus();
		}
		if(frmObj.txtDisplay.value == "")
		{
			flag = 1 ;
			mesg = mesg + "Please enter display name\n" ;
			frmObj.txtDisplay.focus();
		}

		if(flag == 1)
		{
			alert(mesg);
			return false ;
		}
		return true ;
	}
//-->
	function IsActive(ID,Text)
	{

		document.frmPage.hdTime.value=ID;
		document.frmPage.hdStatus.value=Text;
	//	alert(document.frmActive.hdEmployeeID.value + " " + document.frmActive.hdStatus.value);
		document.frmPage.method="Post";
		document.frmPage.action="active_deactivate_time.asp"
		document.frmPage.submit();
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
											<td valign="top" width="1"><img src="images/blank.gif" border="0" width="1" height="1"></td>
											<td valign="top"><img src="images/rightchk.gif" border="0"></td>
											<td valign="middle" bgcolor="#3F89C3" width="20%" height="31" class="txtfnt" >&nbsp;&nbsp;<b>User:  <%=Session("sms_name") & " " & "(" & Session("sms_EmplNbr") & ")" %></b></td>
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
                            <td valign="top" class="txtHeader">&nbsp;<b>Add Activity
                              Time </b><br>
                              <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                            </td>
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                          </tr>
                          <tr>
                            <td valign="top">
                              <form name="frmChecktype" method="post" action="activitytime_submit.asp" onSubmit="return validateForm();">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <tr class="Content" >
                                    <td bgcolor="#9FC4E1" width="40%">
                                      <div align="right">Time :</div>
                                    </td>
                                    <td bgcolor="#B4DAF8">
                                      <div align="left">
                                        <select name="cboTime" id="cboTime" style="width:120px">
                                          <option value="100">Select Time</option>
                                          <%
										  		For i = 0 to 23
										  %>
                                          <option value="<%=i&":00"%>"><%=i&":00"%></option>
                                          <%
												Next
										  %>
                                        </select>
                                      </div>
                                    </td>
								</tr>
                                  <tr class="Content" >
                                    <td bgcolor="#9FC4E1">
                                      <div align="right">Display Name : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8">
                                      <div align="left">
                                        <input type="textbox" name="txtDisplay" value="" maxlength="25" style="width:120px">
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
                                      <input type="hidden" name="hdact_ID" id="hdact_ID" value="<%=lclstr_hdactID%>">
                                    </td>
                                    <td> <span  align=center ><font color="red" size=2>
                                      <%
										if lclint_MesgNo = 1 then
											Response.Write "Check Type Added Successfully"
										elseif lclint_MesgNo = 2 then
											Response.Write "Sorry! Cannot Add Duplicate Check Type"
										elseif lclint_MesgNo = 3 then
											Response.Write "Check Type Updated Successfully"
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
                            <td valign="top">&nbsp;</td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr>
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">
                              <form name="frmPage">
                                <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#FFFFFF">
                                  <tr align="left" class="txtHeader">
                                    <td  colspan="3">&nbsp;<b>Manage Activity Time </b></td>
                                  </tr>
                                  <tr align="center" class="Content">
                                    <td bgcolor="#9FC4E1"><b>Sl.No</b></td>
                                    <td bgcolor="#9FC4E1"><b>Time</b></td>
                                    <td bgcolor="#9FC4E1"><b>Status</b></td>
                                  </tr>
								  <%
								  	set objRs = Server.createObject("ADODB.Recordset")
									sTime = "select act_ID,act_time,act_Isactive,act_Display from tblActivityTime order by act_Time"
									objRs.Open sTime,ConnFMG,3,3
									if Not objRs.EOF then
										i = 1
										color="#9FC4E1"
										While Not objRs.EOF
										if color="#9FC4E1" then
											color="#B4DAF8"
										else
											color="#9FC4E1"
										end if
									if objRs("act_Isactive") = "True" then
										sActive = "<a href=javascript:IsActive('"& objRs("act_ID") &"','Active') class=Content>De-Activate</a>"
									else
										sActive = "<a href=javascript:IsActive('"& objRs("act_ID") &"','De-Active') class=Content>Activate</a>"
									end if

								  %>
                                  <tr class="Content" bgcolor=<%=color%>>
                                    <td align="center"><%=i%></td>
                                    <td align="center"><%=FormatDateTime(objRs("act_Time"),3) & " - " & objRs("act_Display") %></td>
                                    <td align="center"><%=sActive%></td>
                                  </tr>
                                  <%
								  		i = i + 1
										objRs.MoveNext
										Wend
									else
								  %>
                                  <tr class="Content" bgcolor="#B4DAF8" align="center">
                                    <td colspan="3" bgcolor="#FFFFFF">No Records Found
                                    </td>
                                  </tr>
									<%end if%>
                                  <tr class="Content" bgcolor="#B4DAF8" align="left">
                                    <td colspan="3" bgcolor="#FFFFFF">
									 <input type="hidden" name="hdTime" value="">
									  <input type="hidden" name="hdStatus" value="">
									 </td>
                                  </tr>
                                </table>
                              </form>
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

<%@LANGUAGE="VBSCRIPT"%>
<% if  not (session("sms_validated")) = "True" then
	session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
	Response.Redirect "/default.asp"
 	end if
		Session.Timeout=40
%>
<!--#include file="./includes/connect.asp"-->
<html>
<head>
<title>Welcome to our Intranet | Health Check</title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">


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
				  <% elseif Role_ID = "1" then %>
				  <!--#include file="./includes/navigation.asp"-->
				  <% elseif Role_ID = "2" then %>
				  <!--#include file="./includes/navigation-fmg.asp"-->
				  <% elseif Role_ID = "3" or Role_ID = "4" then %>
				  <!--#include file="./includes/navigation-hd.asp"-->
				  <% else %>
				   <!--#include file="./includes/navigation-null.asp"-->
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
                            <td valign="top" class="txtHeader"><b>Health Check
                              for Data Center</b><br>
                              <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                            </td>
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                          </tr>
                          <tr>
                            <td valign="top">
                               <%if Trim(Session("sms_EmplNbr")) = "1247" then%>
                              <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                <tr class="Content" bgcolor="#9FC4E1" >
                                  <td bgcolor="#FFFFFF">
                                    <div align="left"> &nbsp;</div>
                                  </td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">
                                    <div align="left">&nbsp;<a href="manage_add_areas.asp" class="content">Area</a></div>
                                  </td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1"> &nbsp;<a href="manage_admin_checktypes.asp" class="content">Check
                                    Type</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_admin_checkpoints.asp" class="content">Check
                                    Point</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1">&nbsp;<a href="manage_add_alertusers.asp" class="content">Alert
                                    Users</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_admin_users.asp" class="content">Users</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1">&nbsp;<a href="manage_add_roles.asp" class="content">Roles</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_activitytime.asp" class="content">Time</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1">&nbsp;<a href="dashboard_admin.asp" class="content">Location
                                    wise Dashboard</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_reports_location.asp" class="content">Report</a></td>
                                </tr>
                              </table>
                              <br>
							  <%elseif Role_ID = "1" then %>
                              <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                <tr class="Content" bgcolor="#9FC4E1" >
                                  <td bgcolor="#FFFFFF">
                                    <div align="left"> &nbsp;</div>
                                  </td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">
                                    <div align="left">&nbsp;<a href="manage_add_areas.asp" class="content">Area</a></div>
                                  </td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1"> &nbsp;<a href="manage_admin_checktypes.asp" class="content">Check
                                    Type</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_admin_checkpoints.asp" class="content">Check
                                    Point</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1">&nbsp;<a href="manage_add_alertusers.asp" class="content">Alert
                                    Users</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_admin_users.asp" class="content">Users</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1">&nbsp;<a href="manage_add_roles.asp" class="content">Roles</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_activitytime.asp" class="content">Time</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1">&nbsp;<a href="dashboard_admin.asp" class="content">Location
                                    wise Dashboard</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_reports_location.asp" class="content">Report</a></td>
                                </tr>
                              </table>
                              <br>
							  <%elseif Role_ID = "3" or Role_ID = "4" then %>
                              <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                <tr class="Content" bgcolor="#9FC4E1" >
                                  <td bgcolor="#FFFFFF">
                                    <div align="left"> &nbsp;</div>
                                  </td>
                                </tr>
								<% if Role_ID = "3" then%>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">
                                    <div align="left">&nbsp;<a href="helpdesk_activity.asp" class="content">Daily
                                      Activity</a></div>
                                  </td>
                                </tr>
								<% end if %>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1">&nbsp;
								  <% if Role_ID = "4" then%>
								  <a href="dashboard_spc.asp" class="content">My Dashboard</a>
								  <% else %>
								  <a href="dashboard_user.asp" class="content">My Dashboard</a>
								  <% end if %>
								  </td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="dashboard_user_all.asp" class="content">All
                                    Location Dashboard</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1">&nbsp;<a href="dashboard_admin.asp" class="content">Location
                                    wise Dashboard</a></td>
                                </tr>
                              </table>
                              <br>
							  <%elseif Role_ID = "2" then%>
                              <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                <tr class="Content" bgcolor="#9FC4E1" >
                                  <td bgcolor="#FFFFFF">
                                    <div align="left"> &nbsp;</div>
                                  </td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">
                                    <div align="left">&nbsp;<a href="dashboard_user_all.asp" class="content">All
                                      Locations Dashboard</a></div>
                                  </td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#9FC4E1"> &nbsp;<a href="dashboard_admin.asp" class="content">Location
                                    wise Dashboard</a></td>
                                </tr>
                                <tr class="Content" >
                                  <td bgcolor="#B4DAF8">&nbsp;<a href="manage_reports_location.asp" class="content">Report</a></td>
                                </tr>
                              </table>
							  <br>	
							  <% else %>	
								
                              <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                <tr class="Content" bgcolor="#9FC4E1" align="center" > 
                                  <td bgcolor="#FFFFFF"> 
                                    <div align="center"><font color="red"> <b>You dont have 
                                      permission to access this application</b></font></div>
                                  </td>
                                </tr>
                              </table>
							  		
							  <%end if%>
                            </td>
                          </tr>
                          <tr>
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">&nbsp;</td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr>
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">&nbsp; </td>
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

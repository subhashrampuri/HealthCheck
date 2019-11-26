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

<html>
<head>
<title>Welcome to our Intranet | Health Check of Datacenter</title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
<script language="javascript">
	function redirect(ID,LOC,ROLE)
	{
		//alert(ID);
		document.frmUsers.hdEmpID.value=ID;
		document.frmUsers.hdLOC.value=LOC;
		document.frmUsers.cboRole.value=ROLE;
		document.frmUsers.method="Post";
		document.frmUsers.action="manage_employee_details.asp"
		document.frmUsers.submit();
	} 
	function IsActive(ID,Text,LCode)
	{

		document.frmActive.hdEmployeeID.value=ID;
		document.frmActive.hdStatus.value=Text;
		document.frmActive.hdLocation.value=LCode;
	
	//	alert(document.frmActive.hdEmployeeID.value + " " + document.frmActive.hdStatus.value);	
		document.frmActive.method="Post";
		document.frmActive.action="active_deactivate.asp"
		document.frmActive.submit();
	}
</script>		

</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#ffffff">
  <table cellpadding="0" cellspacing="0" border="0"  bgcolor="#ffffff" height="100%">
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
                            <td valign="top" class="txtHeader"><b>Manage Users</b><br>
                              <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                            </td>
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                          </tr>
                          <tr> 
                            <td valign="top"> </td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top"> 
                              <form name="frmUsers">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <tr align="center" class="Content"> 
                                    <td colspan="7" align="right"><a href="manage_search_users.asp" class="Content"><b>Add Users</b></a></td>
                                  </tr>
                                  <tr align="center" class="Content"> 
                                    <td bgcolor="#9FC4E1"><b>Sl. No</b></td>
                                    <td bgcolor="#9FC4E1"><b>Name</b></td>
                                    <td bgcolor="#9FC4E1"><b>Email</b></td>
                                    <td bgcolor="#9FC4E1"><b>Location</b></td>
                                    <td bgcolor="#9FC4E1"><b>User Type</b></td>
                                    <td bgcolor="#9FC4E1"><b>Is Active</b></td>
                                    <td bgcolor="#9FC4E1"><b>Edit</b></td>
                                  </tr>
                              <%
								Set adoRs = Server.CreateObject("ADODB.RecordSet")
								MySql = "select usr_ID,usr_EmpID,usr_Name,usr_Email,rol_ID,loc_ShortCode,dbo.fn_Roles(rol_ID) as rol_Name,dbo.fn_Location(loc_ShortCode) as location,isActive from tblusers"
								adoRs.Open MySql,ConnFMG,3,3
								if adoRs.Eof = false then
								i  = 1
								color="#9FC4E1"

								lclint_MaxPageSize = 10
								adoRs.PageSize = lclint_MaxPageSize
								lclint_MaxPageCount = adoRs.PageCount
								adoRs.AbsolutePage = lclint_currPage
								i= 1		
								lclint_listNum = 0
								While Not (adoRs.Eof or (lclint_listNum+1) > adoRs.PageSize)
								lclint_listNum = lclint_listNum + 1
								
								if adoRs("location") <> "" then
									loc_name = adoRs("location")
									l = 1
								else
									loc_name = ""
									l = 0
								end if
								if color="#9FC4E1" then
									color="#B4DAF8"
								else
									color="#9FC4E1"
								end if	
								
								if adors("isActive") = "True" then
									sActive = "<a href=javascript:IsActive('"& adoRs("usr_EmpID") &"','Active','"& adoRs("loc_ShortCode") &"') class=Content>De-Activate</a>"
								else
									sActive = "<a href=javascript:IsActive('"& adoRs("usr_EmpID") &"','De-Active','"& adoRs("loc_ShortCode") &"') class=Content>Activate</a>"									
								end if
								userID = adors("usr_ID")
								
							%>
                                  <tr class="Content"> 
                                    <td align="center" bgcolor="<%=color%>"><%=i%></td>
                                    <td align="left" bgcolor="<%=color%>"><%=adoRs("usr_name")%></td>
                                    <td align="left" bgcolor="<%=color%>"><%=adoRs("usr_Email")%></td>
                                    <td align="center" bgcolor="<%=color%>"><%=loc_Name%></td>
                                    <td align="center" bgcolor="<%=color%>"><%=adoRs("rol_Name")%></td>
                                    <td align="center" bgcolor="<%=color%>"><%=sActive%></td>
                                    <td align="center" bgcolor="<%=color%>">
									<% if adoRs("rol_ID") = "3" then%>
									<a href="javascript: redirect('<%=adoRs("usr_EmpID")%>','<%=adoRs("loc_ShortCode")%>',<%=adoRs("rol_ID")%>)" class="Content">Edit</a>
									<% end if%>
									</td>
                                  </tr>
                                  <% 
								  	i = i + 1
									adoRs.MoveNext
									wend
								else%>
                                  <tr class="Content" align="center"> 
                                    <td colspan="7" bgcolor="#FFFFFF">No Records 
                                      Found</td>
                                  </tr>
                                  <%	
								end if
									
								%>
                                  <tr class="Content" align="right"> 
                                    <td colspan="7" bgcolor="#FFFFFF"> 
                                      <div align="right">Page : 
                                        <%
										lclint_pageNum = 1
										while lclint_pageNum <= lclint_MaxPageCount
											if CINT(lclint_currPage) = lclint_pageNum then
											Response.Write lclint_currPage
											else
											%>
                                        <A  href="manage_admin_users.asp?pno=<%=lclint_pageNum%>" ><%=lclint_pageNum%></a> 
                                        <%
											end if
											Response.Write "&nbsp;"
										lclint_pageNum = lclint_pageNum+1
										wend
									%>
                                      </div>
                                    </td>
                                  </tr>
                                  <tr class="Content" align="left"> 
                                    <td colspan="7" bgcolor="#FFFFFF">
									<input type="hidden" name="hdEmpID" value="">
									<input type="hidden" name="acttype" value="E">
									<input type="hidden" name="hdLOC" value="">
									<input type="hidden" name="cboRole" value="">
									</td>
                                  </tr>
                                </table>
                              </form>
                            </td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr>
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">
                              <form name="frmActive" >
							  <input type="hidden" name="hdEmployeeID" value="">
							  <input type="hidden" name="hdStatus" value="">
							  <input type="hidden" name="hdLocation" value="">
							  
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

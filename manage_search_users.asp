<%@LANGUAGE="VBSCRIPT"%>
<% 
	if  not (session("sms_validated")) = "True" then
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
	if Request.Form("txtEmpID") <> "" OR Request.Form("txtEmpName") <> "" then
		EmpID=Replace(Server.HTMLEncode(Request.Form("txtEmpID")),"'","''")
		EmpName=Replace(Server.HTMLEncode(Request.Form("txtEmpName")),"'","''")
	end if
%>
<html>
<head>
<title>Welcome to our Intranet | Health Check </title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
<script language="javascript">
	function redirect(ID)
	{
		//alert(ID);
		document.frmPage.hdEmpID.value=ID;
		document.frmPage.method="Post";
		document.frmPage.action="manage_employee_details.asp"
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
                            <td valign="top" class="txtHeader"><b>Search Users</b><br>
                              <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                            </td>
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                          </tr>
                          <tr>
                            <td valign="top">
                              <form name="frmSearch" method="post" action="manage_search_users.asp">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                  <tr class="Content" >
                                    <td colspan="2">Search by Employee Number
                                      or Employee Name</td>
                                  </tr>
                                  <tr class="Content" >
                                    <td bgcolor="#9FC4E1" width="40%">
                                      <div align="right">EmployeeID : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8">
                                      <div align="left">
                                        <input type="text" name="txtEmpId" maxlength="10">
                                      </div>
                                    </td>
                                  </tr>
                                  <tr class="Content" align="center" bgcolor="#B4DAF8" >
                                    <td colspan="2">
                                      <div align="right"></div>
                                      <div align="center">OR</div>
                                    </td>
                                  </tr>
                                  <tr class="Content" >
                                    <td bgcolor="#9FC4E1">
                                      <div align="right">Employee Name: </div>
                                    </td>
                                    <td bgcolor="#B4DAF8">
                                      <div align="left">
                                        <input type="text" name="txtEmpName" maxlength="100">
                                        (<b><font color="#FF0000">Please enter
                                        first name only</font></b>)</div>
                                    </td>
                                  </tr>
                                  <tr class="Content" >
                                    <td bgcolor="#B4DAF8" colspan="2">&nbsp;</td>
                                  </tr>
                                  <tr class="Content" >
                                    <td bgcolor="#9FC4E1">&nbsp;</td>
                                    <td bgcolor="#B4DAF8">
                                      <input type="submit" name="button" value="Search">
                                      <input type="reset" name="Reset" value="Reset">
                                    </td>
                                  </tr>
                                  <tr class="Content" >
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                  </tr>
                                </table>
                              </form>
                            </td>
                          </tr>
                          <tr>
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">
				<%
					'response.write EmpID & "<br>" & EmpName

					Set objRs =  Server.CreateObject("ADODB.Recordset")
					if EmpID <> ""  then

						sQuery = "SELECT Employee.EmplNbr, Employee.EmplName, EmpDetails.email, EmpDetails.Deptcode " & _
							"  FROM Employee INNER JOIN EmpDetails ON Employee.EmplNbr = EmpDetails.Emplnbr " & _
							"  WHERE (EmpDetails.Deptcode = 'FMG' or EmpDetails.Deptcode = '20' or EmpDetails.Deptcode = '10' ) " & _ 
							"  and Employee.EmplNbr = '" & Trim(EmpID) & "' "

					elseif Request.Form("txtEmpName") <> "" then

						sQuery = "SELECT Employee.EmplNbr, Employee.EmplName, EmpDetails.email, EmpDetails.Deptcode " & _
							"  FROM Employee INNER JOIN EmpDetails ON Employee.EmplNbr = EmpDetails.Emplnbr " & _
							"  WHERE (EmpDetails.Deptcode = 'FMG' or EmpDetails.Deptcode = '20' or EmpDetails.Deptcode = '10' ) " & _ 
							"  and Employee.EmplName like '%" & EmpName & "%' "
					else
						sQuery = "SELECT Employee.EmplNbr, Employee.EmplName, EmpDetails.email, EmpDetails.Deptcode " & _
							" FROM Employee INNER JOIN EmpDetails ON Employee.EmplNbr = EmpDetails.Emplnbr " & _
							" WHERE (EmpDetails.Deptcode = 'FMG' or EmpDetails.Deptcode = '20' or EmpDetails.Deptcode = '10' ) " & _ 
							" and Employee.EmplNbr = '" & Trim(EmpID) & "' and Employee.EmplName like '%" & EmpName & "%' "
					end if
					'Response.write sQuery
					objRs.Open sQuery,ConnSMS,3,3


				%>
						</td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr>
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">
                              <form name="frmPage">
                                <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#FFFFFF">
                                  <tr align="center" class="Content">
                                    <td bgcolor="#9FC4E1" width="10%"><b>Sl.No</b></td>
                                    <td bgcolor="#9FC4E1"><b>Employee Name</b></td>
                                    <td bgcolor="#9FC4E1" align="center"><b>Employee
                                      ID</b></td>
                                    <td bgcolor="#9FC4E1"><b>Employee Email</b></td>
                                    <td bgcolor="#9FC4E1"><b>User Info</b></td>
                                  </tr>
                                  <%
								  	if objRs.Eof = false then
									color="#9FC4E1"
									i = 1
									While Not objRs.Eof
									if color="#9FC4E1" then
										color="#B4DAF8"
									else
										color="#9FC4E1"
									end if
									ENo = Trim(objRs("EmplNbr"))
								  %>
                                  <tr class="Content" bgcolor=<%=color%>>
                                    <td align="center"><%=i%></td>
                                    <td align="left"><%=objRs("EmplName")%></td>
                                    <td align="center"><%=Trim(objRs("EmplNbr"))%></td>
                                    <td align="center"><%=objRs("email")%></td>
									<%
										sFetch= "Select usr_EmpID,rol_ID from tblUsers where usr_EmpID='"& ENo &"' "
										Set objFetch = Server.CreateObject("ADODB.RecordSet")
										objFetch.Open sFetch,ConnFMG,3,3
										'Response.write sfetch
										if Not objFetch.EOF then
											'isExist = ""
											if objFetch("rol_ID") = "4" then
												isExist = "Assign Role"
											end if		
										else
											isExist = "Assign Role"
										End if
										
									%>
                                    <td align="center">
									<a href="javascript: redirect('<%=ENo%>')" class="Content"><%=isExist%></a>
									</td>
                                  </tr>
                                  <%
								  	i = i + 1
									objRs.MoveNext
									Wend
									Set objRs = Nothing
									Else
								  %>
                                  <tr class="Content" bgcolor="#B4DAF8" align="center">
                                    <td colspan="5" bgcolor="#FFFFFF">No Records
                                      found</td>
                                  </tr>
                                  <%End if%>
                                  <tr class="Content" bgcolor="#B4DAF8" align="right">
                                    <td colspan="5" bgcolor="#FFFFFF">Page :</td>
                                  </tr>
                                  <tr class="Content" bgcolor="#B4DAF8" align="left">
                                    <td colspan="5" bgcolor="#FFFFFF">
                                      <input type="hidden" name="hdEmpID" value="">
                                      <input type="hidden" name="acttype" value="A">
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

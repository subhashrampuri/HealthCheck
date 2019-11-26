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
lclstr_ActionType = REPLACE(Request.QueryString("actionType"),"'","")
lclint_MesgNo = Request.QueryString("msgno")
lclstr_chktypID = REPLACE(Request.QueryString("chktyp_ID"),"'","")

if UCASE(lclstr_ActionType) = UCASE("UPDATE") then
	lclstr_disable = "READONLY"
	Set adoRs = Server.CreateObject("ADODB.RecordSet")
	lclstr_mysql ="select area_ID,chktyp_Name,isActive from tblChecktypes where chktyp_ID=" & lclstr_chktypID & " "
	'response.write lclstr_mysql
	adoRs.Open lclstr_mysql,ConnFMG,3,3
	if not adoRs.eof then
		AreaID = adoRs("area_ID")
		
		Chktyp_Name = adoRs("chktyp_Name")
		if adoRs("isActive") = "True" then
			Active = "Checked"
		end if
		
	End if
	Set adoRs = NOTHING
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
		var frmObj = document.frmChecktype ;
		var mesg = "" ;
		var flag = 0 ;

		if(frmObj.cboArea.value == "0")
		{
			flag = 1 ;
			mesg = mesg + "Please select area\n" ;
			frmObj.cboArea.focus();
		}
		if(frmObj.txtChecktype.value == "")
		{
			flag = 1 ;
			mesg = mesg + "Checktype is required field\n" ;
			frmObj.txtChecktype.focus();
		}
		else if (isWhitespace(frmObj.txtChecktype.value))
		{
			flag = 1 ;
			mesg = mesg + "Checktype is required field\n" ;
			frmObj.txtChecktype.focus();
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
                            <td valign="top" class="txtHeader"><b>&nbsp;Add Check 
                              Types</b><br>
                              <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                            </td>
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                          </tr>
                          <tr> 
                            <td valign="top"> 
                              <form name="frmChecktype" method="post" action="checktype_submit.asp" onSubmit="return validateForm();">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1" width="40%"> 
                                      <div align="right">Area :</div>
                                    </td>
                                    <td bgcolor="#B4DAF8"> 
                                      <%
										Set objRs=Server.Createobject("ADODB.Recordset")
										MySql= "Select area_ID,area_Name from tblAreas where isActive = 1 order by Area_name"
										objRs.Open MySql,ConnFMG,3,3
									%>
                                      <div align="left"> 
                                        <select name="cboArea" id="cboArea" style="width:180px">
                                          <option value="0">Select Area</option>
                                          <%
												While Not objRs.EOF
												if AreaID = objRs("area_ID") then
													str = "selected"
												else
													str = ""	
												end if
											%>
                                          <option value="<%=objRs("area_ID")%>" <%=str%>><%=objRs("area_Name")%></option>
                                          <%
												objRS.MoveNext
												Wend
												set objRS = Nothing
											%>
                                        </select>
                                      </div>
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1"> 
                                      <div align="right">Check Type : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8"> 
                                      <div align="left"> 
                                        <input type="text" name="txtChecktype" id="txtChecktype" style="width:180px" maxlength="1500" value="<%=Chktyp_Name%>">
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
                                      <input type="hidden" name="hdchktyp_ID" id="hdchktyp_ID" value="<%=lclstr_chktypID%>">
                                    </td>
                                    <td> <span  align=center class="Content"><font color="red"> 
                                      <%
										if lclint_MesgNo = 1 then
											Response.Write "Check Type Added Successfully"
										elseif lclint_MesgNo = 2 then
											Response.Write "Sorry! Cannot Add Duplicate Check Type"
										elseif lclint_MesgNo = 3 then
											Response.Write "Check Type Updated Successfully"
										elseif lclint_MesgNo = 4 then
											Response.Write "Check Type Updated Successfully"
										elseif lclint_MesgNo = 5 then
											Response.Write "Check Type Cannot be De-Activated"
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
                                  <tr align="left" class="Content"> 
                                    <td bgcolor="#FFFFFF" colspan="5" class="txtHeader"><b>&nbsp;Manage 
                                      Check Types</b></td>
                                  </tr>
                                  <tr align="center" class="Content"> 
                                    <td bgcolor="#9FC4E1"><b>Sl.No</b></td>
                                    <td bgcolor="#9FC4E1"><b>Check Type</b></td>
                                    <td bgcolor="#9FC4E1"><b>Area</b></td>
                                    <td bgcolor="#9FC4E1"><b>Is Active</b></td>
                                    <td bgcolor="#9FC4E1"><b>Edit</b></td>
                                  </tr>
                                  <%
									Set ChkTypRs = Server.CreateObject("ADODB.Recordset")
									My_Sql = " select a.chktyp_ID,a.area_ID,a.chktyp_Name,a.isActive,b.area_Name " & _ 
										" from tblChecktypes a inner join tblAreas b  on a.area_ID= b.area_ID order by a.chktyp_Name,b.area_Name"
									ChkTypRs.Open My_Sql,ConnFMG,3,3
									'Response.write My_Sql
									if ChkTypRs.Eof = false then

									lclint_MaxPageSize = 10
									ChkTypRs.PageSize = lclint_MaxPageSize
									lclint_MaxPageCount = ChkTypRs.PageCount
									ChkTypRs.AbsolutePage = lclint_currPage
									
									color="#9FC4E1"
									
									lclint_listNum = 0
									i =1
									While Not (ChkTypRs.EOF or (lclint_listNum+1) > ChkTypRs.PageSize)
										lclint_listNum = lclint_listNum + 1									
					
										if ChkTypRs("isActive") = "True" then
											sActive = "Active"
										else
											sActive = "In Active"
										end if
										if color="#9FC4E1" then
											color="#B4DAF8"
										else
											color="#9FC4E1"
										end if	
								%>
                                  <tr class="Content" bgcolor=<%=color%>> 
                                    <td align="center" valign="top"><%=i%></td>
                                    <td align="left" style="word-break: break-all; width:300px;" valign="top"><%=ChkTypRs("chktyp_Name")%></td>
                                    <td align="left" style="word-break: break-all; width:300px;" valign="top"><%=ChkTypRs("area_Name")%></td>
                                    <td align="center" valign="top"><%=sActive%></td>
                                    <td align="center" valign="top"><a class="content" href="manage_admin_checktypes.asp?chktyp_ID=<%=ChkTypRs("chktyp_ID")%>&actionType=Update">Edit</a></td>
                                  </tr>
                                  <%
								  	i = i + 1
									ChkTypRs.MoveNext
									Wend
									else %>
                                  <tr class="Content" bgcolor="#B4DAF8" align="Center">
                                    <td colspan="5" bgcolor="#FFFFFF">No Records 
                                      Found</td>
                                  </tr>
								<%	
									end if	
								%>
                                  <tr class="Content" bgcolor="#B4DAF8" align="right"> 
                                    <td colspan="5" bgcolor="#FFFFFF"> 
                                      <div align="right">Page : 
                                        <%
										lclint_pageNum = 1
										while lclint_pageNum <= lclint_MaxPageCount
											if CINT(lclint_currPage) = lclint_pageNum then
											Response.Write lclint_currPage
											else
											%>
                                        <A class="content"  href="manage_admin_checktypes.asp?pno=<%=lclint_pageNum%>" ><%=lclint_pageNum%></a> 
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
                              </form>
                              <script language="Javascript">
	function GotoPageNumber(pno,Menu)
	{
	  document.frmPage.pno.value=pno;
	  //document.frmVisa.selListCountry.value=Menu;
	  document.frmPage.action="<%=request.servervariables("SCRIPT_NAME")%>";
	  document.frmPage.method="post";
	  document.frmPage.submit();
	}
	</script>
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

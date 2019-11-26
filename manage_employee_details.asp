<%@LANGUAGE="VBSCRIPT"%>
<% if  not (session("sms_validated")) = "True" then
	session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
	Response.Redirect "/default.asp"
 	end if
		Session.Timeout=40
%>
<!--#include file="./includes/connect.asp"-->
<%
Server.ScriptTimeout = 300
	EmpID = Request.Form("hdEmpID")
	acttype = Request.form("acttype")
	Locate_ID = Request.Form("hdLOC")
'	Response.write Locate_ID
	RoleID = Request.Form("cboRole")
'	Response.write acttype
	Set objRs =  Server.CreateObject("ADODB.Recordset")
	sQuery = "SELECT Employee.EmplNbr, Employee.EmplName, EmpDetails.email, EmpDetails.Deptcode " & _
		" FROM Employee INNER JOIN EmpDetails ON Employee.EmplNbr = EmpDetails.Emplnbr " & _
		" WHERE (EmpDetails.Deptcode = 'FMG' or EmpDetails.Deptcode = '20' or EmpDetails.Deptcode = '10' ) " & _
		" and Employee.EmplNbr = '" & Trim(EmpID) & "' "
	objRs.Open sQuery,ConnSMS,3,3
	'response.write sQuery
	sName = objRs("EmplName")
	sEmail = objRs("email")

%>
<%
	Set objCheck = Server.CreateObject("ADODB.Recordset")
	sCheckUser = " select a.loc_shortCode,b.loc_Name from tblUsers as a inner join tblLocations b on " & _
		" a.loc_ShortCode = b.loc_ShortCode where a.IsActive = 1 and rol_ID = 3 "
	Set objCheck = 	ConnFMG.Execute(sCheckuser)
	If objCheck.EOF = false then
		While Not objCheck.EOf
		if sCode <> "" then sCode = sCode& ","   end if
			sCode = sCode& "'"  & objcheck("loc_shortCode") & "'"

		if locationName <> "" then locationName = locationName & " " & "&" & " " end if
			locationName =  locationName & objCheck("loc_Name")
			objCheck.MoveNext
			Wend
		end if
%>
<%
	Set objCheck = Server.CreateObject("ADODB.Recordset")
	sCheckSPC = " select a.loc_shortCode,b.loc_Name from tblUsers as a inner join tblLocations b on " & _
		" a.loc_ShortCode = b.loc_ShortCode where a.IsActive = 1 and rol_ID = 4 "
	Set objSPC = 	ConnFMG.Execute(sCheckSPC)
	If objSPC.EOF = false then
		While Not objSPC.EOf
		if sSPCCode <> "" then sSPCCode = sSPCCode& ","   end if
			sSPCCode = sSPCCode& "'"  & objSPC("loc_shortCode") & "'"
		
'		if locationName <> "" then locationName = locationName & " " & "&" & " " end if
'			locationName =  locationName & objSPC("loc_Name")
			objSPC.MoveNext
			Wend
	else
		sSPCCode = "NO"
	end if
		'Response.write sSPCCode & ":" 
%>


<html>
<head>
<title>Welcome to our Intranet | Health Check of Datacenter</title>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
<Script language=JavaScript src="./includes/validate.js" type=text/javascript></SCRIPT>
<script language="javascript">
	function reload_data()
	{

		var frmObj = document.frmEmpDetails ;
		frmObj.hdEmpID.value = "<%=EmpID%>"
		frmObj.acttype.value = "<%=acttype%>"
		frmObj.method="Post";
		frmObj.action="manage_employee_details.asp"
		frmObj.submit();

	}
	function Validate_Submit()
	{
		var frmObj = document.frmEmpDetails ;
		var mesg = "" ;
		var x=frmObj.elements.length;
		var cnt=0;

		if(frmObj.cboRole.value == "0")
		{
			alert("Please select Role");
			frmObj.cboRole.focus();
			return false
		}
		if (frmObj.cboRole.value == "4")
		{
			if(frmObj.cboLocation.value == "0")
			{
				alert("Please select location");
				frmObj.cboLocation.focus();
				return false
			}
		}
		if (frmObj.cboRole.value == "3")
		{
		if(frmObj.cboLocation.value == "0")
		{
			alert("Please select location");
			frmObj.cboLocation.focus();
			return false
		}
		frmObj.ItemPoint.value = "";
		frmObj.ItemRadioPoint.value = "";

		iNoOfCheck=0;
		iNoOfRadio=0;

		for(i=0;i<x;i++)
		{
			if(frmObj.elements[i].type=="checkbox")
		  		{
						if (frmObj.elements[i].checked==true)
						{
							iNoOfCheck++;
						}
				}
		}
		if(iNoOfCheck==0)
		{
			alert("please select check point");
			return false;
		}
		for(i=0;i<x;i++)
		{
			if(frmObj.elements[i].type=="radio")
	  		{
				if (frmObj.elements[i].checked)
				{
					iNoOfRadio++;
				}
				}
		}

		if (iNoOfCheck !=iNoOfRadio)
		{
			alert("Please select Yes/No for comments!");
			return false;
		}

		for(i=0;i<x;i++)
		if(frmObj.elements[i].type=="checkbox")
		  if(frmObj.elements[i].checked==true)
			{
			if(frmObj.ItemPoint.value=="")
				frmObj.ItemPoint.value=frmObj.elements[i].value;
			else
				frmObj.ItemPoint.value = frmObj.ItemPoint.value + "|" + frmObj.elements[i].value;
			}
		//	alert(frmObj.ItemPoint.value);
	//	alert("Done")
		frmObj.method="Post";
		frmObj.action="employee_details_submit.asp"
		frmObj.submit();

		}
		else
		{
	//	alert("Done")
		frmObj.method="Post";
		frmObj.action="employee_details_submit.asp"
		frmObj.submit();
		}

	}
	function EnableItems(ID,radObj)
	{
		//alert(ID);
		if (document.getElementById("chkPoint_"+ID).checked == true)
		{
			radObj[0].disabled = false;
			radObj[1].disabled = false;
		}
		else
		{
					radObj[0].checked = false;
					radObj[1].checked = false;
					radObj[0].disabled = true;
					radObj[1].disabled = true;
		}
	}


	function enableradio()
	{
		var frmObj = document.frmEmpDetails ;
		var s=frmObj.elements.length;
		for(i=0;i<s;i++)
		if(frmObj.elements[i].type=="radio")
		  if(frmObj.elements[i].checked==true)
			{
				frmObj.elements[i].disabled = false;
				PID = frmObj.elements[i].name;
				var indexarr = PID.split("_");
				var str = indexarr[1]+"_"+indexarr[2]+"_"+indexarr[3];
				r = document.getElementById("chkPoint_"+str).name;
				document.getElementById("chkPoint_"+str).checked=true;
				document.getElementById("rdoPoint_"+str).disabled=false;
			}

	}

</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#ffffff" onLoad="javascript: enableradio();">
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
                        <form name="frmEmpDetails" method="Post" action="employee_details_submit.asp">
                          <table cellpadding="0" cellspacing="0" border="0" width="100%">
                            <tr>
                              <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                              <td valign="top" class="txtHeader"><b>Manage Employee
                                Details </b><br>
                                <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                              </td>
                              <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                            </tr>
                            <tr>
                              <td valign="top">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <tr align="center" class="Content">
                                    <td bgcolor="#9FC4E1" align="right" width="40%">
                                      <input type="hidden" name="hdEmpID" value="">
                                      Employee ID : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=EmpID%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#B4DAF8">
                                    <td align="right" bgcolor="#9FC4E1">Employee
                                      Name : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=sName%></td>
                                  </tr>
                                  <tr class="Content">
                                    <td align="right" bgcolor="#9FC4E1">Email
                                      : </td>
                                    <td align="left" bgcolor="#B4DAF8"><%=sEmail%></td>
                                  </tr>
                                  <tr class="Content" bgcolor="#9FC4E1">
                                    <td align="right" bgcolor="#9FC4E1">Role :
                                    </td>
                                    <td align="left" bgcolor="#B4DAF8">
									<%if acttype = "E" then%>
                                      <select name="cboRole" style="Width:200" onChange="reload_data(this.value)" disabled>
									 <%else%>
									  <select name="cboRole" style="Width:200" onChange="reload_data(this.value)">
									<%end if%>
                                        <option value="0">Select Role</option>
                                        <%
											Set rolRs = server.CreateObject("ADODB.RecordSet")
											sRole = "Select rol_ID,rol_Name from tblRoles"
											rolRs.Open sRole,ConnFMG,3,3
											While Not rolRs.EOF
												if acttype <> "" then
													if cInt(RoleID) = rolRs("rol_ID") then
														sStr = "selected"
													else
														sStr = ""
													end if
												else
													sStr = ""
												end if

										%>
                                        <option value="<%=Trim(rolRs("rol_ID"))%>" <%=sStr%>><%=Trim(rolRs("rol_Name"))%></option>
                                        <%
											rolRs.MoveNext
											Wend
											set rolRs = Nothing
										%>
                                      </select>
                                    </td>
                                  </tr>

								  <%if RoleID = "4" then %>
                                  <tr class="Content" bgcolor="#9FC4E1">
                                    <td align="right">Location : </td>
                                    <td align="left" bgcolor="#B4DAF8">
									<%
										Set locRs = server.CreateObject("ADODB.RecordSet")
										'if ISNULL(sSPCCode)  or sSPCCode = "NO" then 	
											sLoc = "Select loc_ShortCode,loc_Name as loc_Name from tblLocations order by loc_Name"
										'else 
										'	sLoc = "Select loc_ShortCode, loc_Name  from tblLocations where loc_Shortcode not in  (" & sSPCCode & ")  order by loc_Name"
										'end if 

										locRs.Open sLoc,ConnFMG,3,3
										%>
									  <select name="cboLocation" style="Width:200">
                                        <option value="0">Select Location</option>
                                        <%
										While Not locRs.EOF
											if Trim(Locate_ID) =  Trim(locRs("loc_ShortCode")) then
												sQry = "selected"
											else
												sQry = ""
											end if
										%>
                                        <option value="<%=Trim(locRs("loc_ShortCode"))%>" <%=sQry%>><%=Trim(locRs("loc_Name"))%></option>
                                        <%
											locRs.MoveNext
											Wend
											set locRs = Nothing
										%>
                                      </select>
                                      <%'response.write sLoc%>
                                    </td>
                                  </tr>
								  <%end if%>


								  <%if RoleID = "3" then %>
                                  <tr class="Content" bgcolor="#9FC4E1">
                                    <td align="right">Location : </td>
                                    <td align="left" bgcolor="#B4DAF8">
                                      <select name="cboLocation" style="Width:200">
                                        <option value="0">Select Location</option>
                                        <%
										Set locRs = server.CreateObject("ADODB.RecordSet")
										if acttype <> "E" then
											sLoc = "Select loc_ShortCode, loc_Name  from tblLocations where loc_Shortcode not in  (" & sCode & ")  order by loc_ID"
										else
											sLoc = "Select loc_ShortCode,loc_Name as loc_Name from tblLocations order by loc_ID"
										end if


										locRs.Open sLoc,ConnFMG,3,3
										While Not locRs.EOF
										'if acttype <> "A" then
											if Trim(Locate_ID) =  Trim(locRs("loc_ShortCode")) then
												sQry = "selected"
											else
												sQry = ""
											end if
										'else
										'	sQry = ""
										'end if

										%>
                                        <option value="<%=Trim(locRs("loc_ShortCode"))%>" <%=sQry%>><%=Trim(locRs("loc_Name"))%></option>
                                        <%
											locRs.MoveNext
											Wend
											set locRs = Nothing
										%>
                                      </select>
                                      <%'response.write sLoc%>
                                    </td>
                                  </tr>
								  <%end if%>
                                  <tr class="Content" bgcolor="#FFFFFF" align="left">
                                    <td colspan="2">
                                      <%
									if acttype = "A"  and locationName <> "" then
										'Response.write "<font color=red>" & locationName &" has already mapped for " & objRs("EmplName") & " " & "(" & EmpID & ")" & "</font>"
									end if
									%>
                                    </td>
                                  </tr>
                                  <tr class="Content" bgcolor="#FFFFFF">
                                    <td align="right" colspan="2">
                                      <hr>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                            <tr>
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top">
							  <% if RoleID = "3" then %>
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                                  <tr class="Content" valign="middle" >
                                    <td bgcolor="#FFFFFF" colspan="6" class="txtHeader"><b>User
                                      Rights</b></td>
                                  </tr>
                                  <tr class="Content" valign="middle" >
                                    <td width="9%" bgcolor="#9FC4E1" colspan="2" ><b>Areas</b></td>
                                    <td bgcolor="#9FC4E1" ><b>Check Types</b></td>
                                    <td bgcolor="#9FC4E1" ><b>Check Points</b></td>
                                    <td bgcolor="#9FC4E1" align="center" ><b>Is
                                      Active</b></td>
                                    <td width="15%" bgcolor="#9FC4E1" ><b>Comments
                                      Options</b></td>
                                  </tr>
                                  <%
								  	Set adoRs = Server.CreateObject("ADODB.RecordSet")
									sLoop = "Select a.area_ID,a.area_name, b.chktyp_id, b.chktyp_name, c.chkpnt_id, c.chkpnt_name, " & _
										" (Select t.tra_isComments from tblTransaction as  t where t.chkpnt_id=c.chkpnt_id and t.usr_empID='" & Trim(EmpID) & "' and t.loc_ShortCode='"& Locate_ID &"') hasRadValue " & _
										" from tblAreas as a, tblChecktypes as b, tblCheckpoint as c where a.isActive=1 and b.isActive=1 and c.isActive=1 " & _
										" and b.area_ID=a.area_ID and c.chktyp_id=b.chktyp_id order by a.area_name,b.chktyp_name, c.chkpnt_name "
									'response.write sLoop
									adors.Open sLoop,ConnFMG,3,3
									area = ""
									checktype = ""
									checkpoint = ""
									While Not adoRs.EOF
								  %>
                                  <%if area <> adors("area_name") then
								  	area = adors("area_name")
								  %>
                                  <tr class="Content" valign="top">
                                    <td colspan="6" bgcolor="#B4DAF8"><b><%=area%></b>
                                    </td>
                                  </tr>
                                  <%end if%>
                                  <%
								  	if checktype <> adoRs("chktyp_Name") then
									   checktype = adoRs("chktyp_Name")
								  %>
                                  <tr class="Content" valign="top">
                                    <td width="4%" bgcolor="#B4DAF8">&nbsp;</td>
                                    <td colspan="5" bgcolor="#B4DAF8"><b><%=checktype%></b>
                                    </td>
                                  </tr>
                                  <%end if%>
                                  <%
								  	if checkpoint <> adoRs("chkpnt_name") then
										checkpoint = adoRs("chkpnt_name")
								  %>
                                  <tr class="Content" valign="top">
                                    <td bgcolor="#B4DAF8" colspan="2" >&nbsp;</td>
                                    <td bgcolor="#B4DAF8" colspan="2" ><%=checkpoint%>
                                    </td>
                                    <td bgcolor="#B4DAF8" align="center">
                                      <input type="checkbox" name="chkPoint_<%=adoRs("area_ID")%>_<%=adoRs("chktyp_ID")%>_<%=adoRs("chkpnt_ID")%>" onClick="EnableItems(this.value,document.frmEmpDetails.rdoPoint_<%=adoRs("area_ID")%>_<%=adoRs("chktyp_ID")%>_<%=adoRs("chkpnt_ID")%>);"  value="<%=adoRs("area_ID")%>_<%=adoRs("chktyp_ID")%>_<%=adoRs("chkpnt_ID")%>">
                                      <!--<%=adoRs("area_ID")%>_<%=adoRs("chktyp_ID")%>_<%=adoRs("chkpnt_ID")%>-->
                                    </td>
                                    <%
									if adoRs("hasRadValue") <> "" and acttype = "E" then
										if adoRs("hasRadValue") = "True" then
											sCheck = "CheckedT"
										end if
										if adoRs("hasRadValue") = "False" then
											sCheck = "CheckedF"
										end if
									%>
                                    <td bgcolor="#B4DAF8" width="12%">
                                      <input type="radio" name="rdoPoint_<%=adoRs("area_ID")%>_<%=adoRs("chktyp_ID")%>_<%=adoRs("chkpnt_ID")%>" <%if sCheck = "CheckedT" then%>Checked<%end if%> value="1" <% if sCheck <> "CheckedT" then %>  DISABLED <% end if%>>
                                      Yes
                                      <input type="radio" name="rdoPoint_<%=adoRs("area_ID")%>_<%=adoRs("chktyp_ID")%>_<%=adoRs("chkpnt_ID")%>" <%if sCheck = "CheckedF" then%>Checked<%end if%> value="0" <% if sCheck <> "CheckedT" then %>  DISABLED <% end if%>>
                                      No </td>
                                    <%else%>
                                    <td bgcolor="#B4DAF8" width="12%">
                                      <input type="radio" name="rdoPoint_<%=adoRs("area_ID")%>_<%=adoRs("chktyp_ID")%>_<%=adoRs("chkpnt_ID")%>"  value="1"  DISABLED >
                                      Yes
                                      <input type="radio" name="rdoPoint_<%=adoRs("area_ID")%>_<%=adoRs("chktyp_ID")%>_<%=adoRs("chkpnt_ID")%>"  value="0"  DISABLED >
                                      No </td>
                                    <%end if%>
                                  </tr>
                                  <% end if %>
                                  <%
                                  sCheck=""
								  		adoRs.Movenext
										Wend
										adoRs.Close
										Set adors = nothing
								  %>
                                  <tr class="Content" valign="top">
                                    <td bgcolor="#B4DAF8" colspan="6" >
                                      <input type="hidden" name="ItemPoint" value = "">
                                      <input type="hidden" name="ItemRadioPoint" value = "">
                                 	  <!--<input type="hidden" name="hdEmp_ID" value="<%=EmpID%>">
                                      <input type="hidden" name="hdEmpName" value="<%=sName%>">
                                      <input type="hidden" name="hdEmpMail" value="<%=sEmail%>">
                                      <input type="hidden" name="hdLocation" value="">
                                      <input type="hidden" name="hdRole" value="">-->
                                    </td>
                                  </tr>
                                  <tr class="Content" valign="top">
                                    <td colspan="6" bgcolor="#9FC4E1" align="center" >&nbsp;

                                    </td>
                                  </tr>
                                  <tr class="Content" valign="top">
                                    <td colspan="6" >&nbsp;</td>
                                  </tr>
                                </table>
								<% end if%>
                              </td>
                              <td valign="top" width="5%">&nbsp;</td>
                            </tr>
                            <tr>
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top" align="center" bgcolor="#9FC4E1">
                                 	  <input type="hidden" name="hdEmp_ID" value="<%=EmpID%>">
                                      <input type="hidden" name="hdEmpName" value="<%=sName%>">
                                      <input type="hidden" name="hdEmpMail" value="<%=sEmail%>">
                                      <input type="hidden" name="hdLocation" value="">
                                      <input type="hidden" name="hdRole" value="">
									  <input type="hidden" name="acttype" value="">
                                <input type="Button" name="Button" value="Submit" onClick="Validate_Submit();">
                                      &nbsp; &nbsp;
                                      <input type="Reset" name="Button" value="Reset">
                              </td>
                              <td valign="top" width="5%">&nbsp;</td>
                            </tr>
                            <tr>
                              <td valign="top" width="5%">&nbsp;</td>
                              <td valign="top">&nbsp;</td>
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

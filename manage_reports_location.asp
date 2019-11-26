<%@LANGUAGE="VBSCRIPT"%>
<% if  not (session("sms_validated")) = "True" then
	session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
	Response.Redirect "/default.asp"
 	end if
	Session.Timeout=40
%> 
<!--#include file="./includes/connect.asp"-->
 <%
	FromDate = Request.form("txtFromDate")
	ToDate = Request.form("txtToDate")
	location = Trim(Request.Form("cboLocation"))
	Arry = Split(location,",")
	
	'Response.write FromDate & ":" & ToDate & ":" & location 
	
'	if Day(Now()) < 10 then lclstr_Report_Date = "0"&Day(Now())& "/" else lclstr_Report_Date = Day(Now())& "/" end if
'	if Day(Now()) < 10 then lclstr_Report_Date = lclstr_Report_Date & "0"&Month(Now())& "/" else lclstr_Report_Date = lclstr_Report_Date & Month(Now()) & "/" end if	

	if Day(Now()) < 10 then lclstr_Report_Date =  "0"&Month(Now())& "/" else lclstr_Report_Date =  Month(Now()) & "/" end if	
	if Day(Now()) < 10 then lclstr_Report_Date = lclstr_Report_Date & "0"&Day(Now())& "/" else lclstr_Report_Date = lclstr_Report_Date & Day(Now())& "/" end if
	lclstr_Report_Date = lclstr_Report_Date & Year(Now())
	
	lclstr_From_Date = lclstr_Report_Date
	lclstr_To_Date = lclstr_Report_Date
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
		var frmObj = document.frmreports ;

		if(frmObj.txtFromDate.value == "")
		{
			alert("Please select from date");
			frmObj.txtFromDate.focus();
			return false;
		}
		if(frmObj.txtToDate.value == "")
		{
			alert("Please select to date");
			frmObj.txtToDate.focus();
			return false;
		}
		if(frmObj.cboLocation.value == "")
		{
			alert("Please select location");
			frmObj.cboLocation.focus();
			return false;
		}

		return true ;
	}
//-->
function OpenHolidayDateDiv()
{
	var tempDate = document.frmreports.txtFromDate.value;
		var temp_arr = tempDate.split("/");
	    if (document.frmreports.txtFromDate.value!="")
        {
           var tempDay,tempMonth;
//           if (tempDate.substring(0,2)<10)
//               tempDay = tempDate.substring(1,2);
//          if (tempDate.substring(3,5)<10)
//               tempMonth = tempDate.substring(4,5);
           tempDay = temp_arr[0];
           tempMonth = temp_arr[1];

		   document.frmreports.Calendar1.Day = tempDay;
		   document.frmreports.Calendar1.Month = tempMonth;
		   document.frmreports.Calendar1.Year = temp_arr[2];
		}
		else
		{
		   //document.frmreports.txtFromDate.value= "<%if day(Now()) < 10 then response.write "0" & day(Now()) else day(Now()) end if%>" +"/" + "<%if month(Now()) < 10 then response.write "0" & month(Now()) else response.write month(Now()) end if%>" + "/" + "<%=year(Now())%>"
			document.frmreports.txtFromDate.value= "<%if month(Now()) < 10 then response.write "0" & month(Now()) else response.write month(Now()) end if%>" +"/" + "<%if day(Now()) < 10 then response.write "0" & day(Now()) else day(Now()) end if%>" + "/" + "<%=year(Now())%>"		   
		 }

		if(divFromDate.style.visibility=="hidden")
			divFromDate.style.visibility="visible";
		else
			divFromDate.style.visibility="hidden";
}

function OpenHolidayDateDiv1()
{
	var tempDate = document.frmreports.txtToDate.value;
		var temp_arr1 = tempDate.split("/");
        if (document.frmreports.txtToDate.value!="")
        {
           var tempDay,tempMonth;
//           if (tempDate.substring(0,2)<10)
//              tempDay = tempDate.substring(1,2);
//           if (tempDate.substring(3,5)<10)
//               tempMonth = tempDate.substring(4,5);
           tempDay = temp_arr1[0];	//tempDate.substring(0,2)
           tempMonth = temp_arr1[1];	//tempDate.substring(3,5)

		   document.frmreports.Calendar2.Day = tempDay;
		   document.frmreports.Calendar2.Month = tempMonth;
		   document.frmreports.Calendar2.Year = temp_arr1[2];
		}
		else
		{
		  // document.frmreports.txtToDate.value= "<%if day(Now()) < 10 then response.write "0" & day(Now()) else day(Now()) end if%>" +"/" + "<%if month(Now()) < 10 then response.write "0" & month(Now()) else response.write month(Now()) end if%>" + "/" + "<%=year(Now())%>"
		  document.frmreports.txtToDate.value= "<%if month(Now()) < 10 then response.write "0" & month(Now()) else response.write month(Now()) end if%>" +"/" + "<%if day(Now()) < 10 then response.write "0" & day(Now()) else day(Now()) end if%>" + "/" + "<%=year(Now())%>"		   
		 }

		if(divToDate.style.visibility=="hidden")
			divToDate.style.visibility="visible";
		else
			divToDate.style.visibility="hidden";
}

</script>
<SCRIPT LANGUAGE=javascript FOR=Calendar1 EVENT=AfterUpdate>
<!--
 Calendar1_AfterUpdate()
//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=Calendar2 EVENT=AfterUpdate>
<!--
  Calendar2_AfterUpdate()
//-->
</SCRIPT>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function Calendar1_AfterUpdate() {
	var dayHolder, monthHolder
	if (document.frmreports.Calendar1.Day<10)
		dayHolder = "0" + document.frmreports.Calendar1.Day
	else
		dayHolder = document.frmreports.Calendar1.Day

	if (document.frmreports.Calendar1.Month<10)
		monthHolder = "0" + document.frmreports.Calendar1.Month
	else
		monthHolder = document.frmreports.Calendar1.Month

	//document.frmreports.txtFromDate.value = dayHolder+ "/" + monthHolder + "/" + document.frmreports.Calendar1.Year;
	document.frmreports.txtFromDate.value = monthHolder+ "/" + dayHolder + "/" + document.frmreports.Calendar1.Year;
	divFromDate.style.visibility="hidden";
}
function Calendar2_AfterUpdate() {
	var dayHolder, monthHolder
	if (document.frmreports.Calendar2.Day<10)
		dayHolder = "0" + document.frmreports.Calendar2.Day
	else
		dayHolder = document.frmreports.Calendar2.Day

	if (document.frmreports.Calendar2.Month<10)
		monthHolder = "0" + document.frmreports.Calendar2.Month
	else
		monthHolder = document.frmreports.Calendar2.Month

	//document.frmreports.txtToDate.value = dayHolder+ "/" + monthHolder + "/" + document.frmreports.Calendar2.Year;
	document.frmreports.txtToDate.value = monthHolder+ "/" + dayHolder + "/" + document.frmreports.Calendar2.Year;
	divToDate.style.visibility="hidden";
}

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
                            <td valign="top" class="txtHeader"><b>Manage Location 
                              wise Report </b><br>
                              <img src="./images/blank.gif" border="0" width="1" height="4"><br>
                            </td>
                            <td valign="top" width="5%" rowspan="2"><img src="./images/blank.gif" border="0" width="1" height="1"></td>
                          </tr>
                          <tr> 
                            <td valign="top"> 
                              <form name="frmreports" method="post" action="location_report.asp" onSubmit="return validateForm();">
                                <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" align="center">
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1" width="40%"> 
                                      <div align="right">From Date : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8"> 
                                      <div align="left"> 
                                        <input type="text" name="txtFromDate" value = "<%=FromDate%>" maxlength="30" style="width:180px" readonly>
                                        <a href="#"><img src="images/calendar.gif" width="22" height="17" border="0" onclick="javascript:OpenHolidayDateDiv(); return false;"></a></div>
                                    </td>
                                  </tr>
                                  <DIV id="divFromDate" style="POSITION:ABSOLUTE;TOP:182;LEFT:400;visibility: hidden; margin: 20px" > 
                                    <OBJECT id=Calendar1 style="LEFT: 10px; WIDTH: 260px; TOP: 15px; HEIGHT: 165px" height=165
		width=220 classid="clsid:8E27C92B-1264-101C-8A2F-040224009C02" VIEWASTEXT>
                                      <PARAM NAME="_Version" VALUE="524288">
                                      <PARAM NAME="_ExtentX" VALUE="6482">
                                      <PARAM NAME="_ExtentY" VALUE="4366">
                                      <PARAM NAME="_StockProps" VALUE="1">
                                      <PARAM NAME="BackColor" VALUE="15787475">
                                      <PARAM NAME="Year" VALUE="<%=year(now())%>">
                                      <PARAM NAME="Month" VALUE="<%=month(now())%>">
                                      <PARAM NAME="Day" VALUE="30">
                                      <PARAM NAME="DayLength" VALUE="3">
                                      <PARAM NAME="MonthLength" VALUE="3">
                                      <PARAM NAME="DayFontColor" VALUE="10485760">
                                      <PARAM NAME="FirstDay" VALUE="1">
                                      <PARAM NAME="GridCellEffect" VALUE="0">
                                      <PARAM NAME="GridFontColor" VALUE="10485760">
                                      <PARAM NAME="GridLinesColor" VALUE="-2147483646">
                                      <PARAM NAME="ShowDateSelectors" VALUE="-1">
                                      <PARAM NAME="ShowDays" VALUE="-1">
                                      <PARAM NAME="ShowHorizontalGrid" VALUE="-1">
                                      <PARAM NAME="ShowTitle" VALUE="-1">
                                      <PARAM NAME="ShowVerticalGrid" VALUE="-1">
                                      <PARAM NAME="TitleFontColor" VALUE="16711680">
                                      <PARAM NAME="ValueIsNull" VALUE="-1">
                                    </OBJECT> </div>
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1"> 
                                      <div align="right">To Date : </div>
                                    </td>
                                    <td bgcolor="#B4DAF8"> 
                                      <div align="left"> 
                                        <input type="text" name="txtToDate" value = "<%=ToDate%>" maxlength="30" style="width:180px" readonly>
                                        <a href="#"><img src="images/calendar.gif" width="22" height="17" border="0" onclick="javascript:OpenHolidayDateDiv1(); return false;"></a></div>
                                    </td>
                                  </tr>
                                  <DIV id="divToDate" style="POSITION:ABSOLUTE;TOP:215;LEFT:400;visibility: hidden; margin: 20px" > 
                                    <OBJECT id=Calendar2 style="LEFT: 10px; WIDTH: 260px; TOP: 15px; HEIGHT: 165px" height=165
		width=220 classid="clsid:8E27C92B-1264-101C-8A2F-040224009C02" VIEWASTEXT>
                                      <PARAM NAME="_Version" VALUE="524288">
                                      <PARAM NAME="_ExtentX" VALUE="6482">
                                      <PARAM NAME="_ExtentY" VALUE="4366">
                                      <PARAM NAME="_StockProps" VALUE="1">
                                      <PARAM NAME="BackColor" VALUE="15787475">
                                      <PARAM NAME="Year" VALUE="<%=year(now())%>">
                                      <PARAM NAME="Month" VALUE="<%=month(now())%>">
                                      <PARAM NAME="Day" VALUE="30">
                                      <PARAM NAME="DayLength" VALUE="3">
                                      <PARAM NAME="MonthLength" VALUE="3">
                                      <PARAM NAME="DayFontColor" VALUE="10485760">
                                      <PARAM NAME="FirstDay" VALUE="1">
                                      <PARAM NAME="GridCellEffect" VALUE="0">
                                      <PARAM NAME="GridFontColor" VALUE="10485760">
                                      <PARAM NAME="GridLinesColor" VALUE="-2147483646">
                                      <PARAM NAME="ShowDateSelectors" VALUE="-1">
                                      <PARAM NAME="ShowDays" VALUE="-1">
                                      <PARAM NAME="ShowHorizontalGrid" VALUE="-1">
                                      <PARAM NAME="ShowTitle" VALUE="-1">
                                      <PARAM NAME="ShowVerticalGrid" VALUE="-1">
                                      <PARAM NAME="TitleFontColor" VALUE="16711680">
                                      <PARAM NAME="ValueIsNull" VALUE="-1">
                                    </OBJECT> </div>
                                  <tr class="Content" > 
                                    <td bgcolor="#9FC4E1" align="right" vAlign="top">Location 
                                      : </td>
                                    <td bgcolor="#B4DAF8"> 
                                      <%
										Set objRs=Server.Createobject("ADODB.Recordset")
										MySql= "Select loc_ShortCode,loc_Name from tblLocations where isActive =1 order by loc_ShortCode"
										objRs.Open MySql,ConnFMG,3,3
									 %>
                                      <select name="cboLocation" id="cboLocation" style="width:180px;height=80px" multiple>
                                        <%
												While Not objRs.EOF
												for i = 0 to UBound(Arry)
													if Trim(cStr(arry(i)))  = Trim(cStr(objRs("loc_ShortCode"))) then
														str = "selected"
														Exit For		
													else
														str = ""	
													end if
												Next	
											%>
                                        <option value="<%=objRs("loc_ShortCode")%>" <%=str%>><%=objRs("loc_Name")%></option>
                                        <%
												objRS.MoveNext
												Wend
											%>
                                      </select>
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td align="right" bgcolor="#9FC4E1">&nbsp;</td>
                                    <td bgcolor="#B4DAF8"> &nbsp;&nbsp;&nbsp;&nbsp; 
                                      <input type="submit" name="Submit" value="Submit">
                                      <input type="reset" name="Reset" value="Reset">
                                    </td>
                                  </tr>
                                  <tr class="Content" > 
                                    <td align="right">&nbsp;</td>
                                    <td>&nbsp;</td>
                                  </tr>
                                </table>
                              </form>
                            </td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">&nbsp; </td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">&nbsp; </td>
                            <td valign="top" width="5%">&nbsp;</td>
                          </tr>
                          <tr>
                            <td valign="top" width="5%">&nbsp;</td>
                            <td valign="top">&nbsp;</td>
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

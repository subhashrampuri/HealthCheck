<%@LANGUAGE="VBSCRIPT"%>
	<% if  not (session("sms_validated")) = "True" then
		session("lastURLName")= Request.Servervariables("SCRIPT_NAME")
		Response.Redirect "/default.asp"
		end if
	Session.Timeout=40
		Response.Buffer =True
		Server.ScriptTimeout = 10000
		Response.ContentType = "application/vnd.ms-excel"
	%>
	
<!--#include file="./includes/connect.asp"-->
<%
	FromDate = Request.form("txtFromDate")
	ToDate = Request.form("txtToDate")
	location = Trim(Request.Form("cboLocation"))
	Arry = Split(location,",")
	Response.AddHeader "content-disposition","attachment;filename=HealthCheck_" & FromDate & "_" & ToDate & ".xls"

	'Response.write UBound(Arry)
	'Response.write FromDate & ":" & ToDate & ":" & location
	
	flag=false
	MySql = " Select a.area_ID,b.area_Name,a.chktyp_ID, c.chktyp_Name,a.chkpnt_ID,d.chkpnt_Name,a.loc_ShortCode, " & _ 
		" dbo.fn_Location(a.loc_ShortCode) as loc_Name,a.tdt_Comments,a.tdt_Date,a.tdt_ActivityStatus,e.act_Time,e.tdm_Entryby,a.Usr_EmpID,f.usr_Name, " & _ 
		" e.tdm_EditBy,dbo.GetEmplName(e.tdm_EditBy) as Edit_name,e.tdm_editOn " & _ 
		" from tblTransactionDetails a, tblAreas b, tblChecktypes c, tblCheckpoint d, tblTransactionDetailsMaster e,tblUsers f " & _ 
		" where a.area_ID = b.area_ID and a.chktyp_ID = c.chktyp_ID and a.tdm_ID = e.tdm_ID and a.chkpnt_ID = d.chkpnt_ID " & _ 
		" and a.usr_EmpID = f.Usr_EmpID and (DateDiff(DD,'"& FromDate &"',a.tdt_Date)>=0 and DateDiff(DD,a.tdt_Date,'"&ToDate&"')>=0) and ("
	
	for i = 0 to UBound(Arry)
		if UBound(Arry) = 0 then
			MySql = MySql &  "a.Loc_ShortCode =" & "'" & Trim(arry(i))  & "'"
		else
			if flag=false then
				MySql = MySql & " a.Loc_ShortCode =" & "'" & Trim(arry(i)) & "'" 
				flag=true
			else
				MySql = MySql & " or a.Loc_ShortCode =" & "'" & Trim(arry(i)) & "'" 
			end if
		end if		
	Next
	
	MySql = MySql & " ) order by a.tdt_date"
	
	Set objRs = ConnFMG.Execute(MySql)	
	'Response.write MySql & "<br><br>"
%>
<table width="100%" border="1" cellspacing="1" cellpadding="5">
  <tr > 
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Sl.No</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Area</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Check Type</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Check Point</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Location</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Comments</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Activity Date</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Activity Status</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Activity Time Slot</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">User ID</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">User Name</font></b></td>
    <td><b><font face="Verdana, Arial, Helvetica, sans-serif">Entry By</font></b></td>
    <td><b>Edited By</b></td>
    <td><b>Edited On</b></td>
  </tr>
  <%
	i = 1
	if Not objRs.EOF then
	While Not objRs.Eof
	Response.Flush
		if objRs("tdt_Comments") <> "" then
			Comments = objRs("tdt_Comments")
		else
			Comments = ""
		end if
		
		if objRs("tdt_ActivityStatus") = "1" then
			Status = "Success"
		else
			Status = "Failure"
		end if
		if objRs("tdm_EditBy") <> "" then 
			Editby = objRs("Edit_name") & " - " & objRs("tdm_EditBy")
		else	
			EditBy = ""
		end if

		if objRs("tdm_editOn") <> "" then 
			EditOn = objRs("tdm_editOn") 
		else	
			EditOn = ""
		end if
		
%>
  <tr valign="top"> 
    <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><%=i%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=objRs("area_Name")%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=objRs("chktyp_Name")%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=objRs("chkpnt_Name")%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=objRs("loc_Name")%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=Comments%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=FormatDateTime(objRs("tdt_Date"),2)%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=Status%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=FormatDateTime(objRs("act_Time"),3)%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=objRs("usr_EmpID")%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=objRs("usr_Name")%> 
      </font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=objRs("tdm_Entryby")%> 
      </font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=EditBy%></font></td>
    <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><%=EditOn%></font></td>
  </tr>
  <%
 	i = i + 1
 	objRs.MoveNext
	wend
	objRs.Close
 	else
 %>
  <tr align="center" > 
    <td colspan="14"><b><font face="Verdana, Arial, Helvetica, sans-serif">No 
      Records Found</font></b></td>
  </tr>
  <%

 end if
%>
</table>

<!--#include file="./includes/connect.asp"-->
<%
	Session.Timeout=40
   Server.ScriptTimeout = 300

	EmpID = request.Form("hdEmp_ID")
	EmpName = Replace(Server.HTMLEncode(request.Form("hdEmpName")),"'","''")
	EmpEmail = Replace(Server.HTMLEncode(request.Form("hdEmpMail")),"'","''")
	Location = 	Request.Form("cboLocation")
	Role = Request.Form("cboRole")

	'Location = Request.Form("hdLocation")
	'Role = Request.Form("hdRole")
	'Response.write EmpID & "<br>" & EmpName & "<br>" & EmpEmail & "<br>" & Location & "<br>" & Role & "<br>" & EmpActive & "<br>"
	'CheckPoint =  Request.Form("ItemPoint") 
	
	Set objCheck = Server.CreateObject("ADODB.RecordSet")
	sCheck = "Select usr_EmpID from tblUsers where usr_EmpID= '" & Trim(EmpID) &"' and loc_ShortCode='"& Trim(Location) &"' "
	objCheck.Open sCheck,ConnFMG,3,3
	
	if objCheck.EOF then
		
		sUser = "Insert into tblUsers values('" & Trim(EmpID) & "','"& Trim(EmpName) & "','"& Trim(EmpEmail) &"',"& Role &",'"& Trim(Location) &"',1)"
		'Response.write sUser & "<br>"
		ConnFMG.Execute(sUser)
		
		CheckPoint=split(Request.Form("ItemPoint"),"|")
		For i = 0 to Ubound(CheckPoint)
			Comment =  request.form("rdoPoint_"&CheckPoint(i))
			vCheck = Split(CheckPoint(i),"_")
			
				'Response.write vCheck(0) & "~~" & vCheck(1) & "~~"& vCheck(2) & "~~" & Comment & "<br>"
			sTransact = "Insert into tblTransaction values('" & Trim(EmpID) &"',"& vCheck(0) &","& vCheck(1) &","& vCheck(2)&","& Comment&",'"& Trim(Location) &"')" 
			'response.write sTransact
			ConnFMG.Execute(sTransact)
			
		Next
		
	else

		sDelete = "Delete from tblTransaction where usr_EmpID = '"& Trim(EmpID)&"' and loc_Shortcode = '"& Trim(Location) &"'"
		ConnFMG.Execute(sDelete)
					
		CheckPoint=split(Request.Form("ItemPoint"),"|")
		For i = 0 to Ubound(CheckPoint)
			Comment =  request.form("rdoPoint_"&CheckPoint(i))
			vCheck = Split(CheckPoint(i),"_")
			
				'Response.write vCheck(0) & "~~" & vCheck(1) & "~~"& vCheck(2) & "~~" & Comment & "<br>"
			sTransactUpdate = "Insert into tblTransaction values('" & Trim(EmpID) &"',"& vCheck(0) &","& vCheck(1) &","& vCheck(2)&","& Comment &",'"& Trim(Location) &"')"	
			'response.write sTransactUpdate
			ConnFMG.Execute(sTransactUpdate)
			
		Next
		
	end if		
	
%>
<!--#include file="./includes/connect.asp"-->
<%
	Response.redirect "manage_admin_users.asp"
%>

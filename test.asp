
<!--#include file="./includes/connect.asp"-->
<%
		set rs = server.CreateObject("ADODB.Recordset")	
		sTime = "SELECT Convert(Char(5),getdate(),108) as currenttime"
		rs.Open sTime,ConnFMG,3,3
		dtime = Rs("currenttime")
		a= split(dtime,":")
		hr = a(0)  
		min = a(1)
		con_hr = a(0) * 60
		tot = cInt(con_hr) + cInt(min)
		response.write "tot:::" & tot & "<br>"
		
		set slotRs = server.CreateObject("ADODB.Recordset")	
		sSlot = "SELECT act_Time,act_Code from tblActivityTime"
		slotRs.Open sSlot,ConnFMG,3,3
		While not slotRs.EOF
		
		Response.write  cInt(slotrs("act_Time")) & "<br>"

		
		slotRs.Movenext
		wend
		
%>

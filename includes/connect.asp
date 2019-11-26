<%

set ConnSMS = Server.CreateObject("ADODB.CONNECTION")
ConnSMS.open session("Intra_Conn")


set ConnFMG = Server.CreateObject("ADODB.CONNECTION")
ConnFMG.open session("Fmg_Conn")

function SQLEncode(strText)
   strText = trim(strText)
   if strText <> "" and isNull(strText) = False then
     strText = replace(strText,"'","''")
   end if
   SQLEncode = strText
end function

 sCheckLogin = "Select rol_ID from tblUsers where usr_EmpID = '"&Trim(Session("sms_EmplNbr"))&"' "
 Set RsLogin = ConnFMG.Execute(sCheckLogin)
'  response.write 	   sCheckLogin

if RsLogin.EOF = false then
	Role_ID =  RsLogin("rol_ID")
  else
	Role_ID = ""
  end if
%>
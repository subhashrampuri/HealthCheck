<%
'if  not (session("sms_validated")) = "True" then
'	session("lastURLName")=request.servervariables("SCRIPT_NAME")
'	Response.Redirect "/default.asp"
'end if

%>
<!-- #include file="includes/upload.asp" -->

<%
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)

Dim UploadRequest
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest  RequestBin

gblstr_uplFileValue = UploadRequest.Item("fileDoc").Item("Value")
gblstr_uplFileContent = UploadRequest.Item("fileDoc").Item("ContentType")
gblstr_uplFile = UploadRequest.Item("fileDoc").Item("FileName")

set fso = server.CreateObject("Scripting.FileSystemObject")
gblstr_uplFileName = Year(Date) & Month(Date) & Day(Date)& "_"& Hour(TIME) & Minute(TIME) & Second(TIME) & "_" & fso.GetFileName(gblstr_uplFile)
gblstr_uplFilePath = server.MapPath(".") & "\documents\" & gblstr_uplFileName
gblstr_uplFilePathVer = "documents/" & gblstr_uplFileName

if (fso.FileExists(gblstr_uplFilePath)) then
	ErrMsg = "Error <br> *  A file with same name already Exists!"
	Response.Write ErrMsg & "<br>"
	Response.Write ("Press <-- Back to Upload the file again")
	Response.End
else
	Set MyFile = fso.CreateTextFile(gblstr_uplFilePath)
	For i = 1 to LenB(gblstr_uplFileValue)
		MyFile.Write chr(AscB(MidB(gblstr_uplFileValue, i, 1)))
	Next
	MyFile.Close
%>
<script language="JavaScript">
	window.opener.document.frmActivity.txtAttachment.value = "<%=gblstr_uplFilePathVer%>" ;
	window.close();
</script>
<%
end if

%>
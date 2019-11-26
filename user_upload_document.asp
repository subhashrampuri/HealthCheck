<%
if  not (session("sms_validated")) = "True" then
	session("lastURLName")=request.servervariables("SCRIPT_NAME")
	Response.Redirect "/default.asp"
end if
	Session.Timeout=40
Server.ScriptTimeout = 300
%>

<HTML>
<TITLE>Welcome to our Intranet | Health Check of Datacenter</TITLE>
<link rel="stylesheet" type="text/css" href="./includes/style.css">
<Script Language="Javascript">
function ChkFile()
{
if (document.frmUpload.fileDoc.value == "")
	{
	alert("Please select zip/rar file to upload")
	document.frmUpload.fileDoc.focus()
	return false;
	}
var oas;
try
{
	oas=new ActiveXObject("Scripting.FileSystemObject");
}
catch (e)
{
	alert("ActiveX is not enabled in your browser.");
	return false;
}
var d = document.frmUpload.fileDoc.value;
var e = oas.getFile(d);
var f = e.size;
if (parseInt(f) > 4194304)
{
	alert("upload file reached max size");
	return false;
}
else
{
	lclstr_imagename = document.frmUpload.fileDoc.value
	var lclint_imgext = lclstr_imagename.lastIndexOf(".")
	var lclstr_imageext = lclstr_imagename.substring(parseInt(lclint_imgext)+1)
	lclstr_imageext = lclstr_imageext.toUpperCase()
		if(lclstr_imageext != "ZIP" && lclstr_imageext != "RAR")
		{
		alert("Please upload only zip/rar file")
		document.frmUpload.fileDoc.value = ""
		document.frmUpload.fileDoc.focus()
		return
		}
	else
		{
		document.frmUpload.action = "user_upload_document_add.asp"
		document.frmUpload.submit()
		}
}
}

</Script>

<BODY>
<form name=frmUpload method=post action="user_upload_document_add.asp" enctype="multipart/form-data">
  <table width="95%" border="0" cellspacing="1" cellpadding="5" align=center bgcolor="#FFFFFF">
    <tr bgcolor="#9FC4E1"> 
      <td colspan=3 class="txtHeader" align=center><b>Upload File / Document</b></td>
</tr>
    <tr bgcolor="#B4DAF8" class="Content"> 
      <td  width=30% align=right>Attach File &nbsp;</td>
      <td width=20 ><font color="red">*</font></td>
      <td class="blue"> 
        <input type=file size=30 maxlength=200 name="fileDoc" ></td>
</tr>
    <tr bgcolor="#B4DAF8" class="Content"> 
      <td  width=30% align=right>&nbsp;</td>
      <td width=20 ></td>
      <td class="blue"> 
        <input type=button size=30 name="butSubmit" value="UPLOAD" onClick="ChkFile();"></td>
</tr>
</table>
</form>

</BODY>
</HTML>

<% Option Explicit%><!--#include file="__settings/mainsettings.asp"--><%

  function StripPath(str)
    if(str <> "") or (str = NULL) then
      if(instr(str, "/") <> 0) then
        str = left(str, instrrev(str, "/"))
      end if
      StripPath = str
    else
      StripPath = ""
    end if
  end function

  if(Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) <> true) then Response.Redirect("Imager.asp?Dir=" & Request.QueryString("lbdir") & "&Image=" & Request.QueryString("lbimage") & "&Page=" & Request.QueryString("lbpage"))

  dim objFS, Folder, FolderName, sFile, sDir, Level, sType, Indent
  Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	Response.Buffer = true
	Level = 0
	Indent = 5
	sDir = Request.QueryString("dir")
	sFile = Request.QueryString("file")
	sType = Request.QueryString("type")
	
  function ValidatePath(str)
    do while(InStr(str, "//") > 0)
      str = replace(str,"//", "/")
    loop
    if(right(str, 1) = "/") and not (len(str) <= 1) then
      str = left(str, len(str)-1)
    end if
    ValidatePath = str
  end function
	
  function CharString(chr, length)
    dim str, i	
    str = ""
    for i = 0 to length
      str = str & chr
    next
    CharString = str
  end function

	sub GetDirStructure(str)
	  dim objFolder
		Set objFolder = objFS.GetFolder(Server.MapPath(ValidatePath(strStdDir + str)))
		Level = Level + 1
		for each Folder in objFolder.SubFolders
		  FolderName = Folder.Name
			if(FolderName <> "__settings") and (FolderName <> "__thumbs") and (Foldername <> "__languages") then
				Response.Write("					  " & CharString("&nbsp;", Level*Indent) & " : <a href=""#"" onclick=""window.opener.Move" & sType & "('" & sDir & "','" & sFile & "','" & ValidatePath(str + "/" + FolderName) & "'); self.close();"">" & FolderName & "</a><br />" & vbCrLf)
				Response.Flush
				GetDirStructure(str + "/" + FolderName)
			end if
		next
		Level = Level - 1
		Set objFolder = Nothing
  end sub	
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html>
	<head>
		<title>Folder browser</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
		<link href="<%= strCSS %>" rel="stylesheet" type="text/css" />
		<style type="text/css">
			body { background-image: none; }
		</style>
		<script language="JavaScript">
			function Resize()
			{
				if (parseInt(navigator.appVersion)>3) {
					if (navigator.appName=="Netscape") {
						window.outerWidth = document.getElementById('main').style.width;
						window.outerHeight = document.getElementById('main').style.height; 
					}
					if (navigator.appName.indexOf("Microsoft")!=-1) {
						self.resizeTo(document.getElementById('main').style.width, document.getElementById('main').style.height);
					}
				}
			}
		</script>
	</head>
	
  <body onLoad="//Resize();">
    <table id="main">
      <tr>
        <td>
          <p style="align: center;" class="folderbrowser">Select which folder you wish to move the <%= lcase(sType) %> to...</p>
        </td>
      </tr>
      <tr>
        <td>
          <span class=""folderbrowser"">
<%= Response.Write("					 &nbsp; : <a href=""#"" onclick=""window.opener.Move" & sType & "('" & sDir & "','" & sFile & "','" & ValidatePath("/") & "'); self.close();"">root</a><br />" & vbCrLf) %>
<% GetDirStructure("") %>
          </span>
        </td>
      </tr>
    </table>	
  </body>
</html><% Set objFS = Nothing %>
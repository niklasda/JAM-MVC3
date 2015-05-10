<% Option Explicit %>
<!--#include file="__settings/mainsettings.asp"-->
<!--#include file="__settings/imagesettings.asp"-->
<!--#include file="__settings/password.asp"--><%
'*********************************************************************
'DO NOT EDIT THE CODE BELOW THIS IF YOU AIN'T FAMILIAR WITH ASP CODE!*
'                (Actually, don't touch it anyway :)                 *
'*********************************************************************

  '*******************
  '* START FUNCTIONS *
  '*******************

  function WaZZa(str)

    str = replace(str,"<","&lt;")
    str = replace(str,">","&gt;")

    str = replace(str,"[BOLD]","<b>")
    str = replace(str,"[/BOLD]","</b>")
    str = replace(str,"[B]","<b>")
    str = replace(str,"[/B]","</b>")
    str = replace(str,"[ULINE]","<u>")
    str = replace(str,"[/ULINE]","</u>")
    str = replace(str,"[U]","<u>")
    str = replace(str,"[/U]","</u>")
    str = replace(str,vbCrLf,"<br />")

    WaZZa = str
  end function

  function Shorty(str, ln)
    if (len(str) > ln) then
      str = left(str, ln) & "..."
    end if
    Shorty = str
  end function

  function UnSpacer(str)
    str = replace(str," ","%20")
    UnSpacer = str
  end function

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

  function ValidatePath(str)
    do while(InStr(str, "//") > 0)
      str = replace(str,"//", "/")
    loop
    if(right(str, 1) = "/") and not (len(str) <= 1) then
      str = left(str, len(str)-1)
    end if
    ValidatePath = str
  end function

  function ReadDescription(dir, image)
    dim str
    if objFS.FileExists(Server.Mappath(ValidatePath(strStdDir & dir & "/" & image & ".desc"))) then
      set objTextFile = objFS.GetFile(server.mappath(ValidatePath((strStdDir & dir & "/" & image & ".desc"))))
      set TextStream = objTextFile.OpenAsTextStream(ForReading, -2)
      str = Trim(TextStream.ReadAll)
      TextStream.close
    end if

    ReadDescription = str
  end function
  
  function GetFolderTemplate()
    dim Template, Description
    Description = ReadDescription(strDir, objFile.Name)
    Template = tmplFolder
    Template = Replace(Template, "[Name]", objFile.Name)
    Template = Replace(Template, "[NameLink]", "<a href=""?Dir=" & strDir & "/" & objFile.Name & """>" & objFile.Name & "</a>")
    Template = Replace(Template, "[ImageCount]", GetImageCount(objFile.Name))
    Template = Replace(Template, "[FolderCount]", GetDirCount(objFile.Name))
    Template = Replace(Template, "[Description]", Description)
    Template = Replace(Template, "[Edit]", "<a href=""?Admin=Edit&amp;Dir=" & strDir & "&amp;Image=" & objFile.Name & "&amp;Page=" & Request.QueryString("Page") & """>")
    Template = Replace(Template, "[/Edit]", "</a>")
    Template = Replace(Template, "[Remove]", "<a href=""#"" onClick=""Javascript:RemoveFolder('" & strDir & "', '" & objFile.Name & "')"">")
    Template = Replace(Template, "[/Remove]", "</a>")
    Template = Replace(Template, "[Rename]", "<a href=""#"" onClick=""Javascript:PromptRenameFolder('" & strDir & "', '" & objFile.Name & "')"">")
    Template = Replace(Template, "[/Rename]", "</a>")
    Template = Replace(Template, "[Move]", "<a href=""#"" onClick=""Javascript:OpenMoveFolderWindow('" & strDir & "', '" & objFile.Name & "')"">")
    Template = Replace(Template, "[/Move]", "</a>")
    if(Description = "") and (InStr(Template, "[D]") <> 0) and (InStr(Template, "[/D]") <> 0) then Template = Replace(Template, mid(Template, InStr(Template, "[D]"), InStr(Template, "[/D]")-InStr(Template, "[D]")+len("[/D]")), "")
    if(Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) <> true) and (InStr(Template, "[A]") <> 0) and (InStr(Template, "[/A]") <> 0) then Template = Replace(Template, mid(Template, InStr(Template, "[A]"), InStr(Template, "[/A]")-InStr(Template, "[A]")+len("[/A]")), "")
    Template = Replace(Template, "[D]", "")
    Template = Replace(Template, "[/D]", "")
    Template = Replace(Template, "[A]", "")
    Template = Replace(Template, "[/A]", "")
    Template = Replace(Template, vbCrLf & vbCrLf, "")
    Template = Replace(Template, vbCrLf, "<br />")
    GetFolderTemplate = Template
  end function
  
  function GetFileTemplate()
    dim Template, Description
    Description = ReadDescription(strDir, objFile.Name)
    Template = tmplFile
    Template = Replace(Template, "[Name]", objFile.Name)
    Template = Replace(Template, "[NameLink]", "<a href=""?Dir=" & strDir & "&amp;Image=" & objFile.Name & """>" & objFile.Name & "</a>")
    Template = Replace(Template, "[SizeBytes]", objFile.Size)
    Template = Replace(Template, "[SizeKBytes]", FormatNumber(objFile.Size / 1024), 2)
    Template = Replace(Template, "[SizeMBytes]", FormatNumber(objFile.Size / 1024 / 1024), 2)
    Template = Replace(Template, "[Description]", Description)
    Template = Replace(Template, "[Edit]", "<a href=""?Admin=Edit&amp;Dir=" & strDir & "&amp;Image=" & objFile.Name & "&amp;Page=" & Request.QueryString("Page") & """>")
    Template = Replace(Template, "[/Edit]", "</a>")
    Template = Replace(Template, "[Remove]", "<a href=""#"" onClick=""Javascript:RemoveFile('" & strDir & "', '" & objFile.Name & "')"">")
    Template = Replace(Template, "[/Remove]", "</a>")
    Template = Replace(Template, "[Rename]", "<a href=""#"" onClick=""Javascript:PromptRenameFile('" & strDir & "', '" & objFile.Name & "')"">")
    Template = Replace(Template, "[/Rename]", "</a>")
    Template = Replace(Template, "[Move]", "<a href=""#"" onClick=""Javascript:OpenMoveFileWindow('" & strDir & "', '" & objFile.Name & "')"">")
    Template = Replace(Template, "[/Move]", "</a>")
    if(Description = "") and (InStr(Template, "[D]") <> 0) and (InStr(Template, "[/D]") <> 0) then Template = Replace(Template, mid(Template, InStr(Template, "[D]"), InStr(Template, "[/D]")-InStr(Template, "[D]")+len("[/D]")), "")
    if(Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) <> true) and (InStr(Template, "[A]") <> 0) and (InStr(Template, "[/A]") <> 0) then Template = Replace(Template, mid(Template, InStr(Template, "[A]"), InStr(Template, "[/A]")-InStr(Template, "[A]")+len("[/A]")), "")
    Template = Replace(Template, "[D]", "")
    Template = Replace(Template, "[/D]", "")
    Template = Replace(Template, "[A]", "")
    Template = Replace(Template, "[/A]", "")
    Template = Replace(Template, vbCrLf & vbCrLf, "")
    Template = Replace(Template, vbCrLf, "<br />")
    GetFileTemplate = Template
  end function

  function GetDefaultValue(str, default)
    if(len(str) < 1) then GetDefaultValue = default else GetDefaultValue = str
  end function

  function GetRandomImage(dir)
    dim imgcount, i, img
    Set objFolderRandom = objFS.GetFolder(Server.MapPath(ValidatePath(strStdDir & Request.QueryString("Dir") & "/" & dir)))

    imgcount = 0
    for each objFileRandom in objFolderRandom.Files
      if IsValidImage(objFileRandom.Name) then
        imgcount = imgcount + 1
      end if
    next

    img = int(rnd*imgcount)
    if img = 0 then img = 1

    i = 0
    for each objFileRandom in objFolderRandom.Files
      if IsValidImage(objFileRandom.Name) then
        i = i + 1
        if(i = img) then
          img = UnSpacer(objFileRandom.Name)
        end if
      end if
    next

    GetRandomImage = img
  end function

  function GetImageCount(dir)
    dim count, Folder, File
    Set Folder = objFS.GetFolder(Server.MapPath(ValidatePath(strStdDir + Request.QueryString("Dir") + "/" + dir)))

    count = 0
    for each File in Folder.Files
      if IsValidImage(File.Name) then count = count + 1
    next

    GetImageCount = count
  end function

  function GetDirCount(dir)
    dim count, Folder, SubFolder
    Set Folder = objFS.GetFolder(Server.MapPath(ValidatePath(strStdDir + Request.QueryString("Dir") + "/" + dir)))

    count = 0
    for each SubFolder in Folder.SubFolders
      if IsValidFolder(SubFolder.Name) then count = count + 1
    next

    GetDirCount = count
  end function

  function IsValidFolder(FolderName)
    dim bolIsValid
    bolIsValid = false
    if(FolderName <> "__settings") and (FolderName <> "__thumbs") and (objFile.Name <> "__languages") then bolIsValid = true
    IsValidFolder = bolIsValid
  end function

  function IsValidImage(FileName)
    dim bolIsValid
    bolIsValid = false
		if(instr(lcase(aImageTypes), GetFileExt(lcase(FileName)))) then bolIsValid = true
    IsValidImage = bolIsValid
  end function
	
	function GetFileExt(FileName)
		dim aFileName
		aFileName = split(FileName, ".")
		GetFileExt = aFileName(UBound(aFileName))
	end function

	function GetParentDir()
    Dim str
    str = strDir
    if(str <> strStdDir) then
      if(InStrRev(str, "/") <> 0) then
        str = left(str, instrrev(str, "/")-1)
      else
        str = ""
      end if
    end if

    GetParentDir = str
  end function

  '*****************
  '* END FUNCTIONS *
  '*****************

  '*******************
  '* INITIATE SCRIPT *
  '*******************

  const ForReading = 1, ForWriting = 2

  dim objFile, objFolder, strDir
  dim Dir, Image, ImageCount, PageCount
  dim objFS, objTextFile, TextStream, PrevFile
  dim objFolderRandom, objFileRandom
  dim version
  dim tmplFile, tmplFolder

  version = "5.4"
  Randomize

  strDir = Request.QueryString("Dir")
  
  set objFS = Server.CreateObject ("Scripting.FileSystemObject")
  set objFolder = objFS.GetFolder(Server.MapPath(ValidatePath(strStdDir + strDir)))
  
  Response.Buffer = true

  if(Request.Form("login") = strAdminPass) then
    Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) = true
  end if
  if(Request.QueryString("Action") = "Logout") then Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) = false

  set objTextFile = objFS.GetFile(server.mappath(ValidatePath(("__settings/file.tmpl"))))
  set TextStream = objTextFile.OpenAsTextStream(ForReading, -2)
  tmplFile = TextStream.ReadAll
  TextStream.close
  set objTextFile = objFS.GetFile(server.mappath(ValidatePath(("__settings/folder.tmpl"))))
  set TextStream = objTextFile.OpenAsTextStream(ForReading, -2)
  tmplFolder = TextStream.ReadAll
  TextStream.close
  set objTextFile = Nothing
  set TextStream = Nothing

  ImageCount = 0
  for each objFile in objFolder.Files
    if IsValidImage(objFile.Name) then
      ImageCount = ImageCount + 1
    end if
  next

  PageCount = cint(ImageCount)/cint(strRows*strThumbs)
  if(PageCount = fix(PageCount)) then PageCount = PageCount -1
  PageCount = fix(PageCount)  

%><?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html>
<head>
  <title><%= strTitle %> - Imager Gallery <%= version %></title>
  <meta name="description" content="<%= strTitle %> - Imager Gallery <%= version %>" />
  <meta name="keywords" content="<%= strTitle %> - Imager Gallery <%= version %>" />
  <meta name="generator" content="Imager Gallery <%= version %>" />
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <link href="<%= strCSS %>" rel="stylesheet" type="text/css" />
</head>

<body>
<% if(strLogo <> "") then %>  <p><img src="<%= strLogo %>" alt="" /></p><% end if %>
<%
  if(Request.QueryString("Image") <> "") Then
    if not(objFS.FileExists(Server.Mappath(ValidatePath(strStdDir & "/" & strDir & "/" & Request.QueryString("Image"))))) and (Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) <> true) then Response.Redirect("?Dir=" & Request.QueryString("Dir"))
%>
  <p class="imagebody">
    <a href="?Dir=<%= Request.QueryString("Dir") %>&amp;Page=<%= Request.QueryString("Page")%>">Return to index</a>
  </p>
  <p class="imagebody" style="text-align: center;">
<%
    for each objFile in objFolder.Files
      if IsValidImage(objFile.Name) then

        if ( objFile.Name = Request.QueryString("Image") ) then
          Response.Write("    <br /><a href=""?Dir=" & strDir & "&amp;Image=" & PrevFile & "&amp;Page=" & Request.QueryString("Page") & """>&lt;&lt; Previous</a>")
        end if

        if ( PrevFile = Request.QueryString("Image") ) then
          Response.Write(" - <a href=""?Dir=" & strDir & "&amp;Image=" & objFile.Name & "&amp;Page=" & Request.QueryString("Page") & """>Next &gt;&gt;</a>")
        end if

        PrevFile = objFile.Name
      end if
    next

    Response.Write(vbCrLf)

    if(strFullWidth <> "") then
      Response.Write("    <br /><br /><a href=""" & ValidatePath(strStdDir & strDir) & "/" & Request.QueryString("image") & """><img src=""GetImage.asp?Image=" & ValidatePath(UnSpacer(strStdDir) & "/" & strDir & "/"& Request.QueryString("image")) & "&amp;Width=" & strFullWidth & "&amp;Height=" & strFullHeight & """ alt="""" /></a>" & vbCrLf)
    else 
      Response.Write("    <br /><br /><a href=""" & ValidatePath(strStdDir & strDir) & "/" & Request.QueryString("image") & """><img src=""" & ValidatePath(strStdDir & strDir & "/" & Request.QueryString("image")) & """ alt="""" /></a>" & vbCrLf)
    end if
%>
  </p>
<%
    '* Check if the user is an admin *
    if (Request.QueryString("Admin") = "Edit") and (Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) = true) then
%>
  <form action="FileManager.asp?action=description&amp;Dir=<%= Request.QueryString("Dir")%>&amp;Image=<%= Request.QueryString("Image")%>&amp;lbdir=<%= Request.QueryString("Dir")%>&amp;lbimage=<%= Request.QueryString("Image")%>&amp;lbpage=<%= Request.QueryString("Page")%>" name="Edit" method="POST">
    <p class="imagebody" style="text-align: center;">
      <textarea name="description" rows="8" cols="75"><% 
  if(ReadDescription(Request.QueryString("Dir"), Request.QueryString("Image")) <> "!_None_!") then
    Response.Write(ReadDescription(Request.QueryString("Dir"), Request.QueryString("Image")))
  end if
%></textarea><br /><br />
      <input type="submit" name="submit" value="Submit Changes!">
    </p>
  </form>
<%
    else
      if(ReadDescription(Request.QueryString("Dir"), Request.QueryString("Image")) <> "!_None_!") then
        Response.Write("  <p class=""description"" style=""text-align: center;"">" & vbCrLf & "    <b>Description:</b><br />" & vbCrLf & "    " & WaZZa(ReadDescription(Request.QueryString("Dir"), Request.QueryString("Image"))) & "" & vbCrLf & "  </p>" & vbCrLf)
      end if
    end if
%>
  <p class="description" style="text-align: center;">
    <a href="javascript: window.print();">Print this page!</a>
  </p>
  <p class="imagebody">
    <a href="?Dir=<%= Request.QueryString("Dir") %>&amp;Page=<%= Request.QueryString("Page")%>">Return to index</a>
  </p>
<%

  Else
    Response.Write("  <table>" & vbCrLf)
    if(strDir <> "") then
      Response.Write("    <tr>" & vbCrLf & "      <td class=""info"" style=""font-size: 12px;"" colspan=""" & strDirThumbs & """><a href=""?Dir=" & GetParentDir & """>Up one dir</a></td>" & vbCrLf & "    </tr>" & vbCrLf)
    end if

    Dim strInt, bolNewRow, bolNeedTR
    strInt = 1
    bolNewRow = true
    bolNeedTR = false

    For Each objFile in objFolder.SubFolders
      
      if(IsValidFolder(objFile.Name)) then
        if(bolNewRow = true) then
          Response.Write("    <tr>" & vbCrLf)
          bolNewRow = false
        end if
        Response.Write("      <td valign=""bottom"" align=""center""><a href=""?Dir=" & strDir & "/" & objFile.Name & """><img src=""GetImage.asp?Image=" & ValidatePath(UnSpacer(strStdDir & "/" & strDir & "/" & objFile.Name & "/" & GetRandomImage(objFile.Name))) & "&amp;Width=" & strTnWidth & "&amp;Height=" & strTnHeight & """ alt=""" & objFile.Name & """ /></a><br />" & vbCrLf)

        Response.Write("        <table cellspacing=""0"" cellpadding=""0"" border=""0"" width=""" & strTnWidth & """>" & vbCrLf)
        Response.Write("          <tr>" & vbCrLf)
        Response.Write("            <td align=""left"">" & vbCrLf)
        Response.Write("              " & GetFolderTemplate() & vbCrLf)
        Response.Write("            </td>" & vbCrLf)
        Response.Write("          </tr>" & vbCrLf)
        Response.Write("        </table>" & vbCrLf)
        Response.Write("      </td>" & vbCrLf)
        bolNeedTR = true
        if(cInt(strInt) = cInt(strDirThumbs)) then
          Response.Write("    </tr>" & vbCrLf)
          bolNewRow = true
          bolNeedTR = false
          strInt = 0
        end if
        strInt = cInt(strInt) + 1
      end if
    next
    if(bolNeedTR = true) then
      Response.Write("    </tr>" & vbCrLf)
      Response.Flush
    end if
    Response.Write("  </table>" & vbCrLf)
    Response.Flush
    
    if(cInt(ImageCount) <> 0) then
      Response.Write("  <table border=""1"" cellspacing=""0"" cellpadding=""1"">" & vbCrLf)

      Dim strSkip, strPage, strSkipped, strDone, strRowCount, i, strPageLow, strPageHigh
      strSkip = 0
      strInt = 1
      strRowCount = 1
      strDone = "False"
      strPage = Request.QueryString("page")
      if(strPage = "") then strPage = 0
      bolNeedTR = true

      if(cInt(strPage) <= 5) then
        strPageLow = 0
        strPageHigh = 10
      elseif(cInt(strPage) >= cInt(PageCount - 5)) then
        strPageLow = strPage - 10
        strPageHigh = PageCount
      else
        strPageLow = strPage - 5
        strPageHigh = strPage + 5
      end if

      if(PageCount > 0) then
        if(cInt(strPage) <> 0) then
          Response.Write("    <tr>" & vbCrLf & "      <td colspan=""" & strThumbs & """ align=""right"" style=""border-bottom: 1px solid;""><span class=""pages"">Pages <a href=""?Dir=" & strDir & "&amp;page=" & cInt(strPage - 1) & """>&lt;&lt;</a> ")
        else
          Response.Write("    <tr>" & vbCrLf & "      <td colspan=""" & strThumbs & """ align=""right"" style=""border-bottom: 1px solid;""><span class=""pages"">Pages ")
        end if
        if(PageCount > 10) then
          For i = strPageLow to strPageHigh
            if(cInt(i) = cInt(strPage)) then
              Response.Write(i+1 & " ")
            else
              Response.Write("<a href=""?Dir=" & strDir & "&amp;page=" & i & """>" & i+1 & "</a> ")
            end if
          Next
        else
          For i = 0 to PageCount
            if(cInt(i) = cInt(strPage)) then
              Response.Write("<b>" & i & "</b> ")
            else
              Response.Write("<a href=""?Dir=" & strDir & "&amp;page=" & i & """>" & i & "</a> ")
            end if
          Next
        end if
        if(cInt(strPage) <> cInt(PageCount)) then
          Response.Write("<a href=""?Dir=" & strDir & "&amp;page=" & cInt(strPage + 1) & """>&gt;&gt;</a></span>" & vbCrLf & "      </td>" & vbCrLf & "    </tr>" & vbCrLf)
        else
          Response.Write("      </td>" & vbCrLf & "    </tr>" & vbCrLf)
        end if
      end if
      for each objFile in objFolder.Files
        strSkipped = "False"
        if IsValidImage(objFile.Name) then
          if(Cint(strSkip) < Cint(strPage*strRows*strThumbs)) then
            strSkip = strSkip + 1
            strSkipped = "True"
          end if
        end if

        if (strSkipped = "False") and (strDone = "False") then
          if IsValidImage(objFile.Name) then

            if(cInt(strInt) = 1) then
              Response.Write("    <tr>" & vbCrLf & "      <td valign=""top"" align=""center"">" & vbCrLf)
            else
              Response.Write("      <td valign=""top"" align=""center"">" & vbCrLf)
            end if

            Response.Write("        <a href=""?Dir=" & strDir & "&amp;Image=" & objFile.Name & "&amp;Page=" & Request.QueryString("Page") & """><img src=""GetImage.asp?Image=" & ValidatePath(UnSpacer(strStdDir & "/" & strDir & "/" & objFile.Name)) & "&amp;Width=" & strTnWidth & "&amp;Height=" & strTnHeight & """ alt=""" & objFile.Name & " - " & cInt(objFile.Size/1000) & "kb"" /></a>" & vbCrLf)

            Response.Write("        <table cellspacing=""0"" cellpadding=""0"" border=""0"" width=""" & strTnWidth & """>" & vbCrLf)
            Response.Write("          <tr>" & vbCrLf)
            Response.Write("            <td align=""left"">" & vbCrLf)
            Response.Write("              " & GetFileTemplate() & vbCrLf)
            Response.Write("            </td>" & vbCrLf)
            Response.Write("          </tr>" & vbCrLf)
            Response.Write("        </table>" & vbCrLf)

            if(cInt(strInt) < cInt(strThumbs)) then
              Response.Write("      </td>" & vbCrLf)
              strInt = strInt + 1
              bolNeedTR = true
            else
              Response.Write("      </td>" & vbCrLf & "    </tr>" & vbCrLf)
              If(Cint(strRowCount) = Cint(strRows)) then strDone = "True"
              strRowCount = strRowCount + 1
              strInt = 1
              bolNeedTR = false
            end if
          end if
        end if
      next
      if(bolNeedTR = true) then
        Response.Write("    </tr>" & vbCrLf)
        Response.Flush
      end if
      Response.Write("  </table>" & vbCrLf)
      Response.Flush
    end if
  end if
  if not(Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) = true) then %>  <form method="post" action="">
    <p><span class="login" style="border: 0px;">Admin login:</span> <input type="password" name="login" class="login" />&nbsp;<input type="submit" class="login" /></p>
  </form>
  <% else %>
  <script type="text/javascript">
    function RemoveFile(dir, file)
    {
      var conf = confirm("You really want to remove " + file + "?")
      if (conf == true)
      {
        window.location = "FileManager.asp?action=removefile&dir=" + dir + "&file=" + file + "&lbdir=<%= Request.QueryString("Dir") %>&lbimage=<%= Request.QueryString("Image") %>&lbpage=<%= Request.QueryString("Page") %>"
      }
      else
      {
        alert("Action canceled.")
      }
    }
    function RemoveFolder(dir, file)
    {
      var conf = confirm("You really want to remove " + dir + "? All files in the folder will be removed.")
      if (conf == true)
      {
        window.location = "FileManager.asp?action=removefolder&dir=" + dir + "&file=" + file + "&lbdir=<%= Request.QueryString("Dir") %>&lbimage=<%= Request.QueryString("Image") %>&lbpage=<%= Request.QueryString("Page") %>"
      }
      else
      {
        alert("Action canceled.")
      }
    }
    function PromptCreate(dir)
    {
      var name = prompt("Enter the name of the new folder")
      if (name != null && name != "")
      {
        window.location = "FileManager.asp?action=createfolder&dir=" + dir + "&target=" + name + "&lbdir=<%= Request.QueryString("Dir") %>&lbimage=<%= Request.QueryString("Image") %>&lbpage=<%= Request.QueryString("Page") %>"
      }
      else
      {
        alert("The foldername must be at least 1 char!")
      }
    }
    function PromptRenameFile(dir, file)
    {
      var name = prompt("What do you want to rename " + file + " to?")
      if (name != null && name != "")
      {
        window.location = "FileManager.asp?action=renamefile&dir=" + dir + "&file=" + file + "&target=" + name + "&lbdir=<%= Request.QueryString("Dir") %>&lbimage=<%= Request.QueryString("Image") %>&lbpage=<%= Request.QueryString("Page") %>"
      }
      else
      {
        alert("The new name must be at least 1 char!")
      }
    }
    function PromptRenameFolder(dir, file)
    {
      var name = prompt("What do you want to rename " + file + " to?")
      if (name != null && name != "")
      {
        window.location = "FileManager.asp?action=renamefolder&dir=" + dir + "&file=" + file + "&target=" + name + "&lbdir=<%= Request.QueryString("Dir") %>&lbimage=<%= Request.QueryString("Image") %>&lbpage=<%= Request.QueryString("Page") %>"
      }
      else
      {
        alert("The new name must be at least 1 char!")
      }
    }
		function OpenMoveFolderWindow(dir, file)
		{
			window.open("FolderBrowser.asp?type=Folder&dir=" + dir + "&file=" + file,"","toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=no, copyhistory=no, width=400, height=400");
		}
    function MoveFolder(dir, file, target)
    {
      if (file != null && file != "" && target != null && target != "")
      {
        window.location = "FileManager.asp?action=movefolder&dir=" + dir + "&file=" + file + "&target=" + target + "&lbdir=<%= Request.QueryString("Dir") %>&lbimage=<%= Request.QueryString("Image") %>&lbpage=<%= Request.QueryString("Page") %>"
      }
      else
      {
        alert("An error occured!")
      }
    }
		function OpenMoveFileWindow(dir, file)
		{
			window.open("FolderBrowser.asp?type=File&dir=" + dir + "&file=" + file,"","toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=no, copyhistory=no, width=400, height=400");
		}
    function MoveFile(dir, file, target)
    {
      if (file != null && file != "" && target != null && target != "")
      {
        window.location = "FileManager.asp?action=movefile&dir=" + dir + "&file=" + file + "&target=" + target + "&lbdir=<%= Request.QueryString("Dir") %>&lbimage=<%= Request.QueryString("Image") %>&lbpage=<%= Request.QueryString("Page") %>"
      }
      else
      {
        alert("An error occured!")
      }
    }
  </script>
  <p class="login" style="border: 0px;">
    <b>Admin section</b><br />
    <form>
      <% if(Request.QueryString("Image") = "") then %><input type="button" onClick="Javascript:PromptCreate('<%= strDir %>')" value="Create new folder" /> <% end if %>
      <% if(allowUpload = "true") and (Request.QueryString("Image") = "") then %><input type="button" onClick="Javascript:window.location = '?Dir=<%= strDir %>&amp;Admin=Upload&amp;Type=Upload'" value="Upload Images" /> <% end if %>
      <% if(allowUpload = "true") and (Request.QueryString("Image") = "") then %><input type="button" onClick="Javascript:window.location = '?Dir=<%= strDir %>&amp;Admin=Upload&amp;Type=UploadThumbs'" value="Upload Thumbnails" /><% end if %>
      <% if(Request.QueryString("Image") <> "") then %><input type="button" onClick="Javascript:window.location = '?Dir=<%= strDir %>&amp;Image=<%= Request.QueryString("Image") %>&amp;Page=<%= GetDefaultValue(Request.QueryString("Page"), 0) %>&amp;Admin=Edit'" value="Edit description" />
      <input type="button" onClick="Javascript:RemoveFile('<%= strDir %>', '<%= Request.QueryString("Image") %>')" value="Remove file" /><% end if %>
      <input type="button" onClick="Javascript:window.location = 'setup.asp'" value="Goto setup" />
      <input type="button" onClick="Javascript:window.location = '?Action=Logout'" value="Logout" />
    </form>
  </p>
  
  <% if(allowUpload = "true") and (Request.QueryString("Image") = "") and (Request.QueryString("Admin") = "Upload") then %><form action="UploadFiles.asp?Dir=<%= Request.QueryString("Dir") %>&Type=<%= Request.QueryString("Type") %>" method="post" enctype="multipart/form-data">
    <input type="file" name="file1" id="file1" class="upload" size="50" /><br />
    <input type="file" name="file2" id="file2" class="upload" size="50" /><br />
    <input type="file" name="file3" id="file3" class="upload" size="50" /><br />
    <input type="file" name="file4" id="file4" class="upload" size="50" /><br />
    <input type="file" name="file5" id="file5" class="upload" size="50" /><br />
    <input type="submit" name="submit" id="submit" class="login" value="Upload"><br />
  </form><% end if%>
  <% end if%>
  <p class="using">
    Gallery generated by <a href="http://www.crazybeavers.se/products_imagergallery.asp">Imager Gallery <%= version %></a>!<br />
    <a href="http://validator.w3.org/check/referer"><img src="http://www.w3.org/Icons/valid-xhtml11" alt="Valid XHTML 1.1!" height="31" width="88" /></a>
    <% if(strValidCSS = "true") then %><a href="http://jigsaw.w3.org/css-validator/check/referer"><img style="border:0;width:88px;height:31px" src="http://jigsaw.w3.org/css-validator/images/vcss" alt="Valid CSS!" /></a><% end if%>
  </p>
<%
  set TextStream = nothing
  set objFolder = nothing
  set objFS = nothing
%></body>
</html>
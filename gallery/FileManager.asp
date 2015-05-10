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

  function ValidatePath(str)
    do while(InStr(str, "//") > 0)
      str = replace(str,"//", "/")
    loop
    if(right(str, 1) = "/") and not (len(str) <= 1) then
      str = left(str, len(str)-1)
    end if
    ValidatePath = str
  end function

  function SwedeConverter(str)
    str = replace(str,"å","&aring;")
    str = replace(str,"ä","&auml;")
    str = replace(str,"ö","&ouml;")
    str = replace(str,"Å","&Aring;")
    str = replace(str,"Ä","&Auml;")
    str = replace(str,"Ö","&Ouml;")

    SwedeConverter = str
  end function

  dim Dir, QueryDir
	QueryDir = Request.QueryString("Dir")
  Dir = strStdDir & QueryDir & "/"

  dim fso, file, folder
  set fso = Server.CreateObject("Scripting.FileSystemObject")

  if(Request.QueryString("action") = "createfolder") then
    if not(fso.FolderExists(server.mappath(ValidatePath(Dir & Request.QueryString("target"))))) then  
      set folder = fso.CreateFolder(server.mappath(ValidatePath(Dir & Request.QueryString("target"))))
    end if
  end if

  if(Request.QueryString("action") = "removefolder") then
    if(fso.FolderExists(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))) then  
      fso.DeleteFolder server.mappath(ValidatePath(Dir & Request.QueryString("file"))), true
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))) then  
      fso.DeleteFile server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")), true
    end if
  end if

  if(Request.QueryString("action") = "removefile") then
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))) then  
      fso.DeleteFile server.mappath(ValidatePath(Dir & Request.QueryString("file"))), true
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))) then  
      fso.DeleteFile server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")), true
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & "__thumbs/" & Request.QueryString("file"))))) then  
      fso.DeleteFile server.mappath(ValidatePath(Dir & "__thumbs/" & Request.QueryString("file"))), true
    end if
  end if

  if(Request.QueryString("action") = "renamefile") then
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))) then  
      set file = fso.GetFile(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))
      file.Move(server.mappath(ValidatePath(Dir & Request.QueryString("target"))))
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))) then  
      set file = fso.GetFile(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))
      file.Move(server.mappath(ValidatePath(Dir & Request.QueryString("target") & ".desc")))
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & "__thumbs/" & Request.QueryString("file"))))) then  
      set file = fso.GetFile(server.mappath(ValidatePath(Dir & "__thumbs/" & Request.QueryString("file"))))
      file.Move(server.mappath(ValidatePath(Dir & "__thumbs/" & Request.QueryString("target"))))
    end if
  end if

  if(Request.QueryString("action") = "renamefolder") then
    if(fso.FolderExists(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))) then  
      set folder = fso.GetFolder(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))
      folder.Name = Replace(Request.QueryString("target"), "/", "")
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))) then  
      set file = fso.GetFile(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))
      file.Name = Replace(Request.QueryString("target"), "/", "") & ".desc"
    end if
  end if

  if(Request.QueryString("action") = "movefile") then
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))) then  
      set file = fso.GetFile(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))
      file.Move(server.mappath(ValidatePath(strStdDir & Request.QueryString("target") & "/" & File.Name)))
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))) then  
      set file = fso.GetFile(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))
      file.Move(server.mappath(ValidatePath(strStdDir & Request.QueryString("target") & "/" & File.Name & ".desc")))
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & "__thumbs/" & Request.QueryString("file"))))) then  
      set file = fso.GetFile(server.mappath(ValidatePath(Dir & "__thumbs/" & Request.QueryString("file"))))
      file.Move(server.mappath(ValidatePath(strStdDir & Request.QueryString("target") & "/__thumbs/" & File.Name)))
    end if
  end if

  if(Request.QueryString("action") = "movefolder") then
    if(fso.FolderExists(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))) then  
      set folder = fso.GetFolder(server.mappath(ValidatePath(Dir & Request.QueryString("file"))))
      folder.Move(server.mappath(ValidatePath(strStdDir & Request.QueryString("target") & "/" & Folder.Name)))
    end if
    if(fso.FileExists(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))) then  
      set file = fso.GetFile(server.mappath(ValidatePath(Dir & Request.QueryString("file") & ".desc")))
      file.Move(server.mappath(ValidatePath(strStdDir & Request.QueryString("target") & "/" & File.Name & ".desc")))
    end if
  end if

  if(Request.QueryString("action") = "description") then
    dim description
    description = SwedeConverter(Request.Form("Description"))
		
    set file = fso.CreateTextFile(server.mappath(ValidatePath(Dir & "/" & Request.QueryString("Image")) & ".desc"), true)
    file.WriteLine(description)
    file.close
  end if

  set fso = Nothing
  set file = Nothing
  set folder = Nothing

  Response.Redirect("Imager.asp?Dir=" & Request.QueryString("lbdir") & "&Image=" & Request.QueryString("lbimage") & "&Page=" & Request.QueryString("lbpage"))
%>
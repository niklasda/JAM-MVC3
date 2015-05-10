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
  
  if(allowUpload <> "true") and (Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) <> true) then Repsonse.Redirect("Imager.asp?Dir=" & Request.QueryString("Dir"))

  function ValidatePath(str)
    do while(InStr(str, "//") > 0)
      str = replace(str,"//", "/")
    loop
    if(right(str, 1) = "/") and not (len(str) <= 1) then
      str = left(str, len(str)-1)
    end if
    ValidatePath = str
  end function

  dim Dir
  if(Request.QueryString("Type") = "Upload") then
    Dir = strStdDir & Request.QueryString("Dir") & "/"
  else
    Dir = strStdDir & Request.QueryString("Dir") & "/__thumbs/"
  end if

  dim fs, folder
  set fs = Server.CreateObject("Scripting.FileSystemObject")
  if not(fs.FolderExists(server.mappath(ValidatePath(Dir)))) then  
    set folder = fs.CreateFolder(server.mappath(ValidatePath(Dir)))
  end if
  set folder = nothing
  set fs = nothing
  
  Server.ScriptTimeout = 600

  dim Upload, i, Error

  if(strUploader = "ASPUpload") then

    set Upload = Server.CreateObject("Persits.Upload")
    Upload.SaveVirtual(ValidatePath(Dir))
    set Upload = Nothing

  elseif(strUploader = "CBUpload") then

    set Upload = Server.CreateObject("CrazyBeavers.Upload")
    for i = 0 to Upload.Count -1
      Error = Upload.SaveToFile(i, ValidatePath(Dir), false)
    next
    set Upload = Nothing

  elseif(strUploader = "aspSmartUpload") then

    set Upload = Server.CreateObject("aspSmartUpload.SmartUpload")
    Upload.AllowedFilesList = "jpg,png,gif"
    Upload.Upload
    i = Upload.Save(ValidatePath(Dir))
    set Upload = Nothing

  elseif(strUploader = "Chili!Upload") then

    set Upload = Server.CreateObject("Chili.Upload.1")
    Upload.SaveToFile(Server.Mappath(ValidatePath(Dir & "/" & Upload.SourceFileName)))
    set Upload = Nothing

  elseif(strUploader = "DynuUpload") then

    set Upload = Server.CreateObject("Dynu.Upload")
    Upload.SavePath = Server.Mappath(ValidatePath(Dir & "/"))
    Upload.OverwriteFiles = false
      Error = Upload.Upload()
    set Upload = Nothing
    
  elseif(strUploader = "JSScript") then
    ' THIS PART OF THE SCRIPT WAS WRITTEN BY GEZ LEMON OF JUICY STUDIO
    ' http://www.juicystudio.com
    ' MODIFIED AND SPEEDED UP FOR IMAGER GALLERY BY KARL-JOHAN SJÖGREN
    Dim postedData, binData, counter, contentType, errorMsg
    Dim boundary, formData, uploadRequest
    Dim fso, browserType, startPos, endPos
    Dim filePath, fileName, savePath, savefile, FileCount 
    Dim requestFiles(5, 1)

    binData = Request.BinaryRead(Request.TotalBytes)
    postedData = ByteArrayToString(binData)
    'For counter = 1 To LenB(binData)
    '  postedData = postedData & Chr(AscB(MidB(binData, counter, 1)))
    'Next
    contentType = Request.ServerVariables("HTTP_CONTENT_TYPE")

    If InStr(contentType, "multipart/form-data") > 0 Then
      endPos = InStrRev(contentType, "=")
      boundary = Trim(Right(contentType, Len(contentType) - endPos))
      formData = Split(postedData, boundary)
      Set uploadRequest = CreateObject("Scripting.Dictionary")
      parseFormData
    Else
      errorMsg = "Incorrect encoding type"
    End If

    Set fso = server.createObject("Scripting.FileSystemObject")
    browserType = UCase(Request.ServerVariables("HTTP_USER_AGENT"))

    For counter = 0 To FileCount - 1
      ' Strip the path info out if not a MAC
      If (InStr(browserType, "WIN") > 0) Then
          startPos = InStrRev(requestFiles(counter, 1), "\")
          fileName = Mid(requestFiles(counter, 1), startPos + 1)
      ElseIf (InStr(browserType, "MAC") > 0) Then
          fileName = requestFiles(counter, 1)
      Else
        startPos = InStrRev(requestFiles(counter, 1), "/")
        fileName = Mid(requestFiles(counter, 1), startPos + 1)
      End If
      filePath = ValidatePath(Dir & fileName)
      savePath = Server.MapPath(filePath)
      Set saveFile = fso.CreateTextFile(savePath, True)
      saveFile.Write(requestFiles(counter, 0))
      saveFile.Close
    Next

    Private Function ByteArrayToString(bArray)
      dim oStream
      Set oStream = Server.CreateObject("ADODB.Stream") 
      oStream.type = 1 'adTypeBinary 
      oStream.mode = 3 'adModeReadWrite 
      oStream.open 
      oStream.write bArray 
      oStream.Position = 0 
      oStream.type = 2 'adTypeText 
      oStream.charset = "ascii"
      ByteArrayToString = oStream.ReadText(oStream.size) 
      oStream.Close 
      Set oStream = Nothing 
    End Function

    Private Sub parseFormData()
    Dim counter, endMarker, fieldInfo, fieldValue
    For counter = 0 To UBound(formData)
      endMarker = InStr(formData(counter), vbCrLf & vbCrLf)
      If endMarker > 0 Then
        fieldInfo = Mid(formData(counter), 3, endMarker - 3)
        fieldValue = Mid(formData(counter), endMarker + 4, Len(formData(counter)) - endMarker - 7)
          If (InStr(fieldInfo, "filename=") > 0) Then
            requestFiles(fileCount, 0) = fieldValue
            requestFiles(fileCount, 1) = getFileName(fieldInfo)
            If requestFiles(fileCount, 1) <> "" Then
              fileCount = fileCount + 1
            End If
          Else
            uploadRequest.add getFieldName(fieldInfo), fieldValue
          End If
        End If
      Next
    End Sub

    Private Function getFieldName(ByVal strFileName)
      Dim startPos, endPos, strQuote
      strQuote = Chr(34)
      startPos = InStr(strFileName, "name=")
      endPos = InStr(startPos + 6, strFileName, strQuote & ";")
      If endPos = 0 Then
        endPos = inStr(startPos + 6, strFileName, strQuote)
      End If
      getFieldName = Mid(strFileName, startPos + 6, endPos - (startPos + 6))
    End Function

    Private Function getFileName(ByVal strFileName)
      Dim startPos, endPos, strQuote
      strQuote = Chr(34)
      startPos = InStr(strFileName, "filename=")
      EndPos = InStr(strFileName, strQuote & vbCrLf)
      getFileName = Mid(strFileName, startPos + 10, endPos - (startPos + 10))
    End Function

    Set saveFile = Nothing
    Set fso = Nothing
    Set uploadRequest = Nothing
    ' END OF SCRIPT WRITTEN BY GEZ LEMON OF JUICY STUDIO
    ' http://www.juicystudio.com
    ' MODIFIED AND SPEEDED UP FOR IMAGER GALLERY BY KARL-JOHAN SJÖGREN
  else
  end if

  Response.Redirect("Imager.asp?Dir=" & Request.QueryString("Dir"))

%>
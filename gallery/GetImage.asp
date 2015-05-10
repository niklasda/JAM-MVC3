<% Option Explicit
  Response.Buffer = True
%><!--#include file="__settings/imagesettings.asp"--><% 

  function ValidatePath(str)
    do while(InStr(str, "//") > 0)
      str = replace(str,"//", "/")
    loop
    if(right(str, 1) = "/") and not (len(str) <= 1) then
      str = left(str, len(str)-1)
    end if
    ValidatePath = str
  end function

  dim FileName, Width, Height, fso, Image, aFileName
  FileName = Server.Mappath(ValidatePath(Request.QueryString("Image")))
  Width = Request.QueryString("Width")
  Height = Request.QueryString("Height")

  set fso = Server.CreateObject("Scripting.FileSystemObject")

  if(right(lcase(strImageNoDir), 3) = "jpg") or (right(lcase(strImageNoDir), 4) = "jpeg") then
    Response.ContentType = "image/jpeg"
  elseif(right(lcase(strImageNoDir), 3) = "gif") then
    Response.ContentType = "image/gif"
  elseif(right(lcase(strImageNoDir), 3) = "png") then
    Response.ContentType = "image/png"
  elseif(right(lcase(strImageNoDir), 3) = "bmp") then
    Response.ContentType = "image/bmp"
  else
    Response.End
  end if

  if not fso.FileExists(FileName) then
    FileName = Server.Mappath(ValidatePath(strImageNoDir))
  end if

  if(strResizer = "Imager Resizer") then

'*************** OLD CODE COMMENTED OUT IN 5.3b
'
'    if(right(lcase(FileName), 3) = "jpg") or (right(lcase(FileName), 4) = "jpeg") then
'      Response.ContentType = "image/jpeg"
'    elseif(right(lcase(FileName), 3) = "gif") then
'      Response.ContentType = "image/gif"
'    elseif(right(lcase(FileName), 3) = "png") then
'      Response.ContentType = "image/png"
'    else
'    end if
'
'    dim xml
'    if(MSXML = 4) then
'      Set xml = Server.CreateObject("Microsoft.XMLHTTP") ' * Creates an instance of MSXML4
'      xml.Open "POST", strImagerDLL & "?Image=" & FileName & "&Width=" & Width & "&Height=" & Height, False
'    else
'      Set xml = Server.CreateObject("MSXML2.ServerXMLHTTP") ' * Creates an instance of MSXML3
'      xml.Open "POST", strImagerDLL & "?Image=" & FileName & "&Width=" & Width & "&Height=" & Height
'    end if

'    xml.Send()
'    Response.BinaryWrite(xml.responseBody)
'    Set xml = Nothing

    Response.Redirect(strImagerDLL & "?Image=" & FileName & "&Width=" & Width & "&Height=" & Height & "&Compression=" & intCompression)

  elseif(strResizer = "PHP Resizer") then

    Response.Redirect("PHPThumb.php?Image=" & FileName & "&Width=" & Width & "&Height=" & Height & "&Compression=" & intCompression)

  elseif(strResizer = "ASP.Net Resizer") then

    Response.Redirect("ASPNetThumb.aspx?Image=" & FileName & "&Width=" & Width & "&Height=" & Height & "&Compression=" & intCompression)

  elseif(strResizer = "ASPJpeg") then

    Set Image = Server.CreateObject("Persits.Jpeg")
    Image.Open(FileName)
    if(Image.OriginalWidth < Image.OriginalHeight) then
      Image.Width = Height
      Image.Height = Width
    else
      Image.Width = Width
      Image.Height = Height
    end if

    Image.Quality = intCompression
    Image.SendBinary
    Set Image = nothing

	elseif(strResizer = "ASPImage") then

    Set Image = Server.CreateObject("ASPImage.Image") 
    if Image.LoadImage(FileName) then
      if Image.MaxX < Image.MaxY  then 
        Image.ResizeR Height, Width
      else 
        Image.ResizeR Width, Height
      end if

			Image.ImageFormat = 1 'JPG
      Image.PixelFormat = 6 '24bit
			Image.JPEGQuality = intCompression
      Response.ContentType = "image/jpeg"
			Response.BinaryWrite(Image.Image)
      Set Image = nothing 
		end if

	elseif(strResizer = "aspSmartImage") then

    Set Image = Server.CreateObject("AspSmartImage.SmartImage")
    Image.OpenFile(FileName)
    
    if Image.OriginalWidth < Image.OriginalHeight  then 
      Image.Resample Height, Width
    else 
      Image.Resample Width, Height
    end if

    aFileName = split(FileName, "/")
    aFileName = split(aFileName(UBound(aFileName)), ".")
		Image.Quality = intCompression
    Image.Download  aFileName(UBound(aFileName)-1) & ".jpg", "image/jpeg", "inline"
    Set Image = nothing 

	elseif(strResizer = "ASPThumb") then

    Set Image = Server.CreateObject("briz.AspThumb")
    Image.Load(FileName)
    
    if Image.Width < Image.Height then 
      Image.Resize Height, Width
    else 
      Image.Resize Width, Height
    end if

		Image.EncodingQuality = intCompression
    Image.Send
    Set Image = nothing

	elseif(strResizer = "csImageFile") then

    Set Image = Server.CreateObject("csImageFile.Manage")
    Image.ReadFile(FileName)
    
    if Image.Width < Image.Height then 
      Image.Resize Height, Width
    else 
      Image.Resize Width, Height
    end if

		Image.JpegQuality = intCompression
    Response.ContentType = "image/jpeg"
    Reponse.BinaryWrite(Image.JPGData)
    Set Image = nothing
    
	elseif(strResizer = "ZImage") then

    Set Image = Server.CreateObject("zimage.zimage") 
    Image.inFile = FileName
    if Image.inWidth < Image.inHeight  then 
      Image.outWidth = Height
      Image.outHeight = Width
    else 
      Image.outWidth = Width
      Image.outHeight = Height
    end if

    Image.outResizeType = 1 'Resize by width and height
    Image.outJpgQuality = intCompression
    Image.outFilter = 6
    Image.outFile = ".jpg"
		Response.ContentType = "image/jpeg"
		Response.BinaryWrite(Image.MakeStream)
    Set Image = nothing 

	elseif(strResizer = "File Thumbnails (Windows)") then

    dim File, FileSize, BinaryStream, Block, Count
    aFileName = split(FileName, "\")
    ReDim Preserve aFileName(UBound(aFileName)+1)
    aFileName(UBound(aFileName)) = aFileName(UBound(aFileName)-1)
    aFileName(UBound(aFileName)-1) = "__thumbs"
    FileName = Join(aFileName, "\")

    if not fso.FileExists(FileName) then FileName = Server.Mappath(ValidatePath(strImageNoDirThumb))

    Set File = fso.GetFile(FileName)
    FileSize = File.Size
    Block = 1024
    
    Set BinaryStream = Server.CreateObject("ADODB.Stream")
    BinaryStream.Open
    BinaryStream.Type = 1
    BinaryStream.LoadFromFile(FileName)

    while FileSize > Block + Count
      Count = Count + Block
      Response.BinaryWrite BinaryStream.Read(Block)
      Response.Flush
    wend

    Response.BinaryWrite BinaryStream.Read(FileSize - Count)
    Response.Flush

    Set BinaryStream = Nothing

	elseif(strResizer = "File Thumbnails (Linux)") then

    aFileName = split(FileName, "/")
    ReDim Preserve aFileName(UBound(aFileName)+1)
    aFileName(UBound(aFileName)) = aFileName(UBound(aFileName)-1)
    aFileName(UBound(aFileName)-1) = "__thumbs"
    FileName = Join(aFileName, "/")

    if not fso.FileExists(FileName) then
      Response.Redirect ValidatePath(strImageNoDirThumb)
    else
      FileName = Request.QueryString("Image")
      aFileName = split(FileName, "/")
      ReDim Preserve aFileName(UBound(aFileName)+1)
      aFileName(UBound(aFileName)) = aFileName(UBound(aFileName)-1)
      aFileName(UBound(aFileName)-1) = "__thumbs"
      FileName = Join(aFileName, "/")
      Response.Redirect(FileName)
    end if

	else
  end if
  
  set fso = nothing
%>
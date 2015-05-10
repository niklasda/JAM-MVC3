<% Option Explicit

  const MainSettings = 1, ImageSettings = 2, Password = 3

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

  '*****************
  '* SAVE SETTINGS *
  '*****************

  if(Request.QueryString("Save") = "Stuff") and (Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) = true) then

    dim objFile, objFS, PassChange
    set objFS = Server.CreateObject ("Scripting.FileSystemObject")

    set objFile = objFS.CreateTextFile(server.mappath("__settings/mainsettings.asp"), true)
    objFile.WriteLine("<%")
    objFile.WriteLine("dim strStdDir, strLogo, strCSS, strValidCSS, strTitle, strTnWidth, strTnHeight, strFullWidth, strFullHeight, strThumbs, strDirThumbs, strRows, strUploader, AllowUpload")
    objFile.WriteLine("strStdDir = """ & Request.Form("strStdDir") & """")
    objFile.WriteLine("strLogo = """ & Request.Form("strLogo") & """")
    objFile.WriteLine("strTitle = """ & Request.Form("strTitle") & """")
    objFile.WriteLine("strTnWidth = """ & Request.Form("strTnWidth") & """")
    objFile.WriteLine("strTnHeight = """ & Request.Form("strTnHeight") & """")
    objFile.WriteLine("strFullWidth = """ & Request.Form("strFullWidth") & """")
    objFile.WriteLine("strFullHeight = """ & Request.Form("strFullHeight") & """")
    objFile.WriteLine("strThumbs = """ & Request.Form("strThumbs") & """")
    objFile.WriteLine("strDirThumbs = """ & Request.Form("strDirThumbs") & """")
    objFile.WriteLine("strRows = """ & Request.Form("strRows") & """")
    objFile.WriteLine("strCSS = """ & Request.Form("strCSS") & """")
    objFile.WriteLine("strValidCSS = """ & Request.Form("strValidCSS") & """")
    objFile.WriteLine("strUploader = """ & Request.Form("strUploader") & """")
    objFile.WriteLine("AllowUpload = """ & Request.Form("AllowUpload") & """")
    objFile.WriteLine("%\>")
    objFile.close

    set objFile = objFS.CreateTextFile(server.mappath("__settings/imagesettings.asp"), true)
    objFile.WriteLine("<%")
    objFile.WriteLine("dim strImagerDLL, strImageNoDir, strImageNoDirThumb, strResizer, aImageTypes, intCompression")
    objFile.WriteLine("strImagerDLL = """ & Request.Form("strImagerDLL") & """")
    objFile.WriteLine("strImageNoDir = """ & Request.Form("strImageNoDir") & """")
    objFile.WriteLine("strImageNoDirThumb = """ & Request.Form("strImageNoDirThumb") & """")
    objFile.WriteLine("strResizer = """ & Request.Form("strResizer") & """")
		objFile.WriteLine("aImageTypes = """ & Request.Form("aImageTypes") & """")
		objFile.WriteLine("intCompression = """ & Request.Form("intCompression") & """")
    objFile.WriteLine("%\>")
    objFile.close
    
    if(Request.Form("oldPassword") <> "") and (Request.Form("oldPassword") = GetSetting("strAdminPass",Password,"")) then
      set objFile = objFS.CreateTextFile(server.mappath("__settings/password.asp"), true)
      objFile.WriteLine("<%")
      objFile.WriteLine("dim strAdminPass")
      objFile.WriteLine("strAdminPass = """ & Request.Form("newPassword") & """")
      objFile.WriteLine("%\>")
      objFile.close
      PassChange = true
    end if

    set objFS = Nothing    
  end if

  '*****************
  '* READ SETTINGS *
  '*****************

  function GetSetting(var, file, default)

    Dim objFile, objFS, stream, line, value
    Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
		value = ""

    if(file = 1) and (objFS.FileExists(server.mappath("__settings/mainsettings.asp"))) then
      Set objFile = objFS.GetFile(server.mappath("__settings/mainsettings.asp"))
      Set stream = objFile.OpenAsTextStream(1, -2)
      do while not stream.atendofstream
        line = stream.ReadLine
        if(var = left(line, len(var))) then
          line = right(line, len(line)-len(var)-4)
          line = left(line, len(line)-1)
          if(line = "") then line = default
          value = line
        end if
      loop
      stream.close
    elseif(file = 2) and (objFS.FileExists(server.mappath("__settings/imagesettings.asp"))) then
      Set objFile = objFS.GetFile(server.mappath("__settings/imagesettings.asp"))
      Set stream = objFile.OpenAsTextStream(1, -2)
      do while not stream.atendofstream
        line = stream.ReadLine
        if(var & " " = left(line, len(var)+1)) then
          line = right(line, len(line)-len(var)-4)
          line = left(line, len(line)-1)
          if(line = "") then line = default
          value = line
        end if
      loop
      stream.close
    elseif(file = 3) and (objFS.FileExists(server.mappath("__settings/password.asp"))) then
      Set objFile = objFS.GetFile(server.mappath("__settings/password.asp"))
      Set stream = objFile.OpenAsTextStream(1, -2)
      do while not stream.atendofstream
        line = stream.ReadLine
        if(var = left(line, len(var))) then
          line = right(line, len(line)-len(var)-4)
          line = left(line, len(line)-1)
          if(line = "") then line = default
          value = line
        end if
      loop
      stream.close
    else
    end if
		if(value = "") then
			GetSetting = default
		else
			GetSetting = value
		end if
  end function

  if(Request.Form("login") <> "") then
    if(Request.Form("login") = GetSetting("strAdminPass",3,"admin")) then
      Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) = true
    end if
  end if
 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
  <title>Imager Gallery - Admin Section</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<style type="text/css">
<!--
td {
  border: 1px solid #000000;
  padding: 5px;
}
p {
  font-size: 11px;
  font-family: Arial;
}
table {
  font-size: 11px;
  font-family: Arial;
  margin: 10px;
  padding: 10px;
  border: 1px solid #000000;
  width: 550px;
}
input.text,select {
  font-size: 11px;
  font-family: Arial;
  border: 1px solid #000000;
  width: 200px;
}
-->
</style>
<script language="javascript">
	function ResizerCheck(x) 
	{
		if(x.value == "Imager Resizer") {
			document.setup.strImagerDLL.disabled=false;
		} else{
			document.setup.strImagerDLL.disabled=true;
		}
		if(x.value == "File Thumbnails (Windows)" || x.value == "File Thumbnails (Linux)"){
			document.setup.strImageNoDirThumb.disabled=false;
		} else{
			document.setup.strImageNoDirThumb.disabled=true;
		}
	}
	function UploaderCheck(x) 
	{
		if(x.checked) {
			document.setup.strUploader.disabled=false;
		} else {
			document.setup.strUploader.disabled=true;
		}
	}
	function CompressionCheck(x) 
	{
		if(parseInt(x.value) < 1 || parseInt(x.value) > 100)
		{
			alert("Compression value must be between 1 and 100!");
			document.setup.intCompression.focus();
		}
	}
</script>
<body<% if(Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) = true) then %> onLoad="ResizerCheck(document.setup.strResizer); UploaderCheck(document.setup.AllowUpload)"<% end if %>>

<p align="center">
  <font size="3">
    Imager Gallery - Admin Section
    <% if(PassChange = true) then Response.Write("<br />Password changed.")%>
  </font>
</p>
<% if(Session("ImagerGalleryAdmin - " & StripPath(Request.ServerVariables("SCRIPT_NAME"))) <> true) then %>
  <form action="setup.asp" method="post" name="setup">
    <table align="center">
      <tr>
        <td>
          Please identify yourself as a administrator of this script by entering the password.
        </td>
        <td>
          <input class="text" type="password" name="login" value="">
        </td>
      </tr>
    </table>
    <p align="center">
      <input class="text" type="submit" value="Login">
    </p>
  </form>
<% else %>
  <p align="center">
    Back to <a href="Imager.asp">Imager Gallery</a>
  </p>
  <form action="setup.asp?Save=Stuff" method="post" name="setup">
    <table align="center">
      <tr>
        <td>
          Root dir of your images. Keep images in subdirs to this folder. Remember to use an ending /, else the edit function won't work.
        </td>
        <td>
          <input class="text" type="text" name="strStdDir" value="<%= GetSetting("strStdDir",MainSettings,left(Request.ServerVariables("PATH_INFO"), len(Request.ServerVariables("PATH_INFO"))-9)) %>">
        </td>
      </tr>
      <tr>
        <td>
            Select which thumbnailer you are going to use.
        </td>
        <td> 
          <select name="strResizer" id="strResizer" onChange="ResizerCheck(this);">
            <option value="ASP.Net Resizer"<% if(GetSetting("strResizer",ImageSettings,"0") = "ASP.Net Resizer") then Response.Write(" selected=""selected""") %>>ASP.Net Resizer</option>
            <option value="ASPImage"<% if(GetSetting("strResizer",ImageSettings,"0") = "ASPImage") then Response.Write(" selected=""selected""") %>>ASPImage</option>
            <option value="ASPJpeg"<% if(GetSetting("strResizer",ImageSettings,"0") = "ASPJpeg") then Response.Write(" selected=""selected""") %>>ASPJpeg</option>
            <option value="aspSmartImage"<% if(GetSetting("strResizer",ImageSettings,"0") = "aspSmartImage") then Response.Write(" selected=""selected""") %>>aspSmartImage</option>
            <option value="ASPThumb"<% if(GetSetting("strResizer",ImageSettings,"0") = "ASPThumb") then Response.Write(" selected=""selected""") %>>ASPThumb</option>
            <option value="csImageFile"<% if(GetSetting("strResizer",ImageSettings,"0") = "csImageFile") then Response.Write(" selected=""selected""") %>>csImageFile</option>
            <option value="Imager Resizer"<% if(GetSetting("strResizer",ImageSettings,"0") = "Imager Resizer") then Response.Write(" selected=""selected""") %>>Imager Resizer</option>
            <option value="File Thumbnails (Windows)"<% if(GetSetting("strResizer",ImageSettings,"0") = "File Thumbnails (Windows)") then Response.Write(" selected=""selected""") %>>File Thumbnails (Windows)</option>
            <option value="File Thumbnails (Linux)"<% if(GetSetting("strResizer",ImageSettings,"0") = "File Thumbnails (Linux)") then Response.Write(" selected=""selected""") %>>File Thumbnails (Linux)</option>
            <option value="PHP Resizer"<% if(GetSetting("strResizer",ImageSettings,"0") = "PHP Resizer") then Response.Write(" selected=""selected""") %>>PHP Resizer</option>
            <option value="zImage"<% if(GetSetting("strResizer",ImageSettings,"0") = "zImage") then Response.Write(" selected=""selected""") %>>zImage</option>
          </select>
        </td>
      </tr>
      <tr>
        <td>
          Specify the web path to the Imager Resizer dll file. This should be in your cgi-bin dir or any other dir that allows scripts and executables to be executed. You can't use a relative link here (i.e /cgi.bin/Imager.dll). Note that it has to be located on the same server as the script even though that you have to specify the full webpath to it.
        </td>
        <td>
          <input class="text" type="text" name="strImagerDLL" value="<%= GetSetting("strImagerDLL",ImageSettings,"http://" & Request.ServerVariables("SERVER_NAME") & "/cgi-bin/Imager.dll")%>">
        </td>
      </tr>
      <tr>
        <td>
          Specify the amount of compression of the generated thumbnails. Lowest value is 1 (low quality, fast download) and the highest is 100 (excellent quality, slow downloads).
        </td>
        <td>
          <input class="text" type="text" name="intCompression" value="<%= GetSetting("intCompression",ImageSettings,"80") %>" onBlur="CompressionCheck(this);">
        </td>
      </tr>
      <tr>
        <td>
            Select which imagetypes that will be shown in the gallery. <a href="javascript:void(0)" onclick="window.open('http://www.crazybeavers.se/Documentation/Imager_Gallery/faq.html#thumbs')">Click here</a> to see which imageformats that are supported with the different components.
        </td>
        <td> 
          <input type="checkbox" name="aImageTypes" id="aImageTypes" value="bmp"<% if(instr(GetSetting("aImageTypes",ImageSettings,""),"bmp")) then Response.Write(" checked=""checked""") %>> BMP<br />
          <input type="checkbox" name="aImageTypes" id="aImageTypes" value="gif"<% if(instr(GetSetting("aImageTypes",ImageSettings,""),"gif")) then Response.Write(" checked=""checked""") %>> GIF<br />
          <input type="checkbox" name="aImageTypes" id="aImageTypes" value="jpg"<% if(instr(GetSetting("aImageTypes",ImageSettings,""), "jpg")) then Response.Write(" checked=""checked""") %>> JPG<br />
          <input type="checkbox" name="aImageTypes" id="aImageTypes" value="jpeg"<% if(instr(GetSetting("aImageTypes",ImageSettings,""), "jpeg")) then Response.Write(" checked=""checked""") %>> JPEG<br />
          <input type="checkbox" name="aImageTypes" id="aImageTypes" value="png"<% if(instr(GetSetting("aImageTypes",ImageSettings,""), "png")) then Response.Write(" checked=""checked""") %>> PNG<br />
        </td>
      </tr>
    </table>
    <br />
    <table align="center">
      <tr>
        <td>
          Check this box if you want to use Imager Gallery's upload functions. This is only avaible for those who have the admin password.
        </td>
        <td align="center">
          Allow upload?<br /><input type="checkbox" name="AllowUpload" value="true"<% if(GetSetting("AllowUpload",1,"false") = "true") then Response.Write(" checked=""checked""") %> onClick="UploaderCheck(this)">
        </td>
      </tr>
      <tr>
        <td>
            Select which upload method you want to use.
        </td>
        <td> 
          <select name="strUploader">
            <option value="aspSmartUpload"<% if(GetSetting("strUploader",MainSettings,"0") = "aspSmartUpload") then Response.Write(" selected=""selected""") %>>aspSmartUpload</option>
            <option value="ASPUpload"<% if(GetSetting("strUploader",MainSettings,"0") = "ASPUpload") then Response.Write(" selected=""selected""") %>>ASPUpload</option>
            <option value="CBUpload"<% if(GetSetting("strUploader",MainSettings,"0") = "CBUpload") then Response.Write(" selected=""selected""") %>>Crazy Beavers Upload</option>
            <option value="Chili!Upload"<% if(GetSetting("strUploader",MainSettings,"0") = "Chili!Upload") then Response.Write(" selected=""selected""") %>>Chili!Upload</option>
            <option value="Dynu Upload"<% if(GetSetting("strUploader",MainSettings,"0") = "Dynu Upload") then Response.Write(" selected=""selected""") %>>Dynu Upload</option>
            <option value="JSScript"<% if(GetSetting("strUploader",MainSettings,"0") = "JSScript") then Response.Write(" selected=""selected""") %>>Juicy Studio Script Uploader</option>
          </select>
        </td>
      </tr>
    </table>
    <br />
    <table align="center">
      <tr>
        <td>
          Page name. This will be showed in the titelbar along with current Imager version.
        </td>
        <td>
          <input class="text" type="text" name="strTitle" value="<%= GetSetting("strTitle",MainSettings,Request.ServerVariables("SERVER_NAME")) %>">
        </td>
      </tr>
      <tr>
        <td>
          Path to image to use as logo on the page. Leave blank to disable.
        </td>
        <td>
          <input class="text" type="text" name="strLogo" value="<%= GetSetting("strLogo",MainSettings,"") %>">
        </td>
      </tr>
      <tr>
        <td>
          Path to CSS Style sheet to be used, all page colors are to be set using css.
        </td>
        <td>
          <input class="text" type="text" name="strCSS" value="<%= GetSetting("strCSS",MainSettings,"style.css") %>">
        </td>
      </tr>
      <tr>
        <td>
          Check this box if your css passes W3C's validation located at <a href="http://jigsaw.w3.org/css-validator/" target="_blank">http://jigsaw.w3.org/css-validator/</a>. This option will display the "Valid CSS" icon on the bottom of the page. The standard Imager Gallery stylesheet validates!
        </td>
        <td align="center">
          My CSS is valid<br /><input type="checkbox" name="strValidCSS" value="true"<% if(GetSetting("strValidCSS",MainSettings,"false") = "true") then Response.Write(" checked=""checked""") %>>
        </td>
      </tr>
      <tr>
        <td>
          Path to JPEG/GIF/PNG image to for empty folders and erroneous files. This <b>can</b> be relative to Imager.asp but it <b>don't have to be</b>.
        </td>
        <td>
          <input class="text" type="text" name="strImageNoDir" value="<%= GetSetting("strImageNoDir",ImageSettings,"__settings/nodir.jpg") %>">
        </td>
      </tr>
      <tr>
        <td>
          Thumbnail of the picture specified above. Only needed when using "File Thumbnails".
        </td>
        <td>
          <input class="text" type="text" name="strImageNoDirThumb" value="<%= GetSetting("strImageNoDirThumb",ImageSettings,"__settings/nodir_thumb.jpg") %>">
        </td>
      </tr>
    </table>
    <br />
    <table align="center">
      <tr>
        <td>
          This setting specifies how many rows of image thumbnails that will be shown on the page.
        </td>
        <td>
          <input class="text" type="text" name="strRows" value="<%= GetSetting("strRows",MainSettings,"3") %>">
        </td>
      </tr>
      <tr>
        <td>
          This setting specifies how many image thumbnails that will be shown on each row.
        </td>
        <td>
          <input class="text" type="text" name="strThumbs" value="<%= GetSetting("strThumbs",MainSettings,"3") %>">
        </td>
      </tr>
      <tr>
        <td>
          This setting specifies how many directory thumbnails that will be shown on each row.
        </td>
        <td>
          <input class="text" type="text" name="strDirThumbs" value="<%= GetSetting("strDirThumbs",MainSettings,"3") %>">
        </td>
      </tr>
      <tr>
        <td>
          Here you define the size of the thumbnails, a simple rule to go by is that the height should be Width/1,33 if the images use standard proportions.
        </td>
        <td>
          Width:<br /><input class="text" type="text" name="strTnWidth" value="<%= GetSetting("strTnWidth",MainSettings,"150") %>"><br />
          Height:<br /><input class="text" type="text" name="strTnHeight" value="<%= GetSetting("strTnHeight",MainSettings,"112") %>">
        </td>
      </tr>
      <tr>
        <td>
          This settings this will set a maximum image size for the full view of the image. Setting them blank will disable this function.
        </td>
        <td>
          Width:<br /><input class="text" type="text" name="strFullWidth" value="<%= GetSetting("strFullWidth",MainSettings,"") %>"><br />
          Height:<br /><input class="text" type="text" name="strFullHeight" value="<%= GetSetting("strFullHeight",MainSettings,"") %>">
        </td>
      </tr>
    </table>
    <br />
    <table align="center">
      <tr>
        <td>
          Here you can change the admin password for the script. You really should change this the first time you set the script up since anyone could read the docs and get the default password.<br />Please remember that the password is <b>case-sensitive</b>!
        </td>
        <td>
          Old password:<br /><input class="text" type="text" name="oldPassword" value=""><br />
          New password:<br /><input class="text" type="text" name="newPassword" value="">
        </td>
      </tr>
    </table>
    <p align="center">
      <input class="text" type="submit" value="Save settings">
    </p>
  </form>
  <p align="center">
    Back to <a href="Imager.asp">Imager Gallery</a>
  </p>
<% end if %>
</body>
</html>
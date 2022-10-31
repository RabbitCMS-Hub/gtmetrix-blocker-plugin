<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************

Class gtmetrix_blocker_plugin
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME, PLUGIN_AUTOLOAD

	Private GTMetrixIP, GTMetrixIPLastFetch ,BotFound, SessionExit, HeaderExit
	Private API_BASE, API_KEY, API_VIEW_PORT, API_WIDTH, API_SS_URL, PLUGIN_STATUS
	Private WEB_SITE_SCREENSHOT, SCREENSHOT_EXPIREDAY, IPLIST_UPDATE_INTERVAL
    Private GTMetrixIPCollection, ShowHTMLResult, ShowPreviewPage
	Private exDB, OLUSTURULDU, SUPER_CACHE_FILE_SUBFIX
	Private GTMETRIX_FETCH_IP_URL

	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"

		' Check Register
		'------------------------------
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE)
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "gtmetrix_blocker_plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "845")
		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "0")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)

		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page [SHOW:UPDATE:AJAX:DELETE]
		'--------------------------------------------------------
		If Query.Data("Page") = "SHOW:FakeGTMetrixPage" Then
			' Call PluginPage("Header")

    		BotFound = True
    		Check()
    		BotFound = True
    		BotPageHTML()
    		BotFound = False

			' Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "AJAX:UpdateGTMetrixIPList" Then
			Call PluginPage("Header")

    		ShowHTMLResult = True
    		
    		GTMetrixIPListUpdate()
    		
    		ShowHTMLResult = False

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_API_CAPTURE_URL", "screenshotlayer.com API Capture URL", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_API_KEY", "screenshotlayer.com API Key", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-3 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_VIEWPORT_RESOLUTION", "Viewport", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-3 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_VIEWPORT_WIDTH", "Viewport Genişlik", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-3 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_SCREENSHOT_EXPIREDAY", "Screenshot Expire Day", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-3 col-sm-12"">"
			.Write  		QuickSettings("input", ""& PLUGIN_CODE &"_IPLIST_UPDATE_INTERVAL", "GTMetrix IP Expire Day", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("select", ""& PLUGIN_CODE &"_SITE_TITLE", "Sahte İçerik Title", "0#Dakika|1#Saat|2#Gün", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("fileajax", ""& PLUGIN_CODE &"_SITE_LOGO", "Sahte İçerik Logo", "/content/files/other-files/", TO_DB)
			.Write "    </div>"
			.Write "</div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <p class=""alert alert-info cms-style"">GTMetrix Test Sunucularını tespit ederek cevap olarak web sayfanızın optimize edilmiş bir ekran görüntüsünü web sitenizmiş gibi sunan akıllı bir eklentidir. Test Sunucu IP adreslerini güncellemeniz önemldir!</p>"
			.Write "    </div>"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=AJAX:UpdateGTMetrixIPList"" class=""btn btn-sm btn-primary"">"
			.Write "        	IP Adreslerini Güncele"
			.Write "        </a>"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:FakeGTMetrixPage"" class=""btn btn-sm btn-danger"">"
			.Write "        	Sahte İçerik Önizleme"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"

			.Write "<div class=""row mt-3"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write 			QuickSettings("tag", "GTMETRIX_BLOCKER_IPLIST", "GTMetrix Test Sunucuları IP Listesi", "", TO_DB)
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
    Private Sub Class_Initialize()
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_CODE  			= "GTMETRIX_BLOCKER"
    	PLUGIN_NAME 			= "GTMetrix Blocker"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/gtmetrix-blocker-plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-devices-off"
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_CREDITS 			= "@badursun Anthony Burak DURSUN"
    	PLUGIN_FOLDER_NAME 		= "Whatsapp-Widget-Plugin"
    	PLUGIN_DB_NAME 			= "aws_log" ' tbl_plugin_XXXXXXX
    	PLUGIN_AUTOLOAD 		= True
    	PLUGIN_ROOT 			= PLUGIN_DIST_FOLDER_PATH(This)
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------

    	GTMETRIX_FETCH_IP_URL 	= "https://gtmetrix.com/locations.html"
    	GTMetrixIPCollection 	= Array()
    	PLUGIN_STATUS 			= Cint( GetSettings("GTMETRIX_BLOCKER_ACTIVE", "0") )
    	' If PLUGIN_STATUS = 0 Then Exit Sub
    	WEB_SITE_TITLE 			= GetSettings(""& PLUGIN_CODE &"_SITE_TITLE", "RabbitCMS")
    	WEB_SITE_LOGO 			= GetSettings(""& PLUGIN_CODE &"_SITE_LOGO", "/content/logo.svg")
    	WEB_SITE_SCREENSHOT 	= "/content/block-screen-shots.jpg"
    	API_BASE 				= GetSettings(""& PLUGIN_CODE &"_API_CAPTURE_URL", "http://api.screenshotlayer.com/api/capture")
    	API_KEY  			 	= GetSettings(""& PLUGIN_CODE &"_API_KEY", "bbbf03c4b9894308864655770a0ab7a8")
    	API_VIEW_PORT 		 	= GetSettings(""& PLUGIN_CODE &"_VIEWPORT_RESOLUTION", "1920x1080")
    	API_WIDTH 			 	= GetSettings(""& PLUGIN_CODE &"_VIEWPORT_WIDTH", "1920")
    	SCREENSHOT_EXPIREDAY 	= Cint( GetSettings(""& PLUGIN_CODE &"_SCREENSHOT_EXPIREDAY", "5") )
    	IPLIST_UPDATE_INTERVAL 	= Cint( GetSettings(""& PLUGIN_CODE &"_IPLIST_UPDATE_INTERVAL", "15") )
    	API_SS_URL 				= DOMAIN_URL
    	BotFound 				= False
    	SessionExit 			= False
    	HeaderExit 				= False
    	ShowHTMLResult 			= False
    	ShowPreviewPage 		= False
    	GTMetrixIPLastFetch 	= GetSettings(""& PLUGIN_CODE &"_IPLIST_FETCHDATE", "26.10.2022 10:00:00")
    	GTMetrixIP 				= GetSettings(""& PLUGIN_CODE &"_IPLIST", "24.109.190.162,172.255.48.130,172.255.48.131")
    	If Instr(1, GTMetrixIP, ",") <> 0 Then
    		GTMetrixIP = Split(GTMetrixIP, ",")
    	Else 
    		GTMetrixIP = Array(GTMetrixIP)
    	End If

    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Register App
    	'-------------------------------------------------------------------------------------
    	class_register()

    	'-------------------------------------------------------------------------------------
    	' Hook Auto Load Plugin
    	'-------------------------------------------------------------------------------------
    	If PLUGIN_STATUS = 1 AND PLUGIN_AUTOLOAD_AT("WEB") = True Then 
    		GTMetrixIPListUpdate()
		    Check()
		    BotPageHTML()
		End If
    End Sub
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------

    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
    Private Sub Class_Terminate()
    	If BotFound=True Then
    		' Response.End
    		Call SystemTeardown("destroy")
    	End If
    End Sub
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------


	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property
	Public Property Get PluginAutoload() 	: PluginAutoload = PLUGIN_AUTOLOAD 			: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable, PluginAutoload)
	End Property
	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------


    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
    Private Property Get ScreenShot()
    	If SSExist(WEB_SITE_SCREENSHOT) = False Then 
    		FetchScreenShot()
		Else
			Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
			Dim objDosya : Set objDosya = objFSO.GetFile(Server.Mappath(WEB_SITE_SCREENSHOT))
				Dim ssDate
					ssDate = objDosya.DateLastModified
				
				If (DateDiff("d", CDate(ssDate), Now()) > SCREENSHOT_EXPIREDAY) Then
    				FetchScreenShot()
				End If
			Set objDosya = nothing
			Set objFSO = nothing
		End If
    End Property
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------

    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
    Private Property Get FetchScreenShot()
        Dim objHTTP : Set objHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0") 
	        With objHTTP
				.Open "GET", API_BASE &"?access_key="&API_KEY&"&url="&API_SS_URL&"&viewport="&API_VIEW_PORT&"&width="&API_WIDTH&""
            	.setRequestHeader "User-Agent", PLUGIN_USER_AGENT
				.setTimeouts 80000, 80000, 80000, 80000
				.Send
			End With

			If objHTTP.Status = 200 Then
				Set Jpeg = Server.CreateObject("Persits.Jpeg")
					Jpeg.OpenBinary( objHTTP.responseBody )
					Jpeg.Save Server.Mappath(WEB_SITE_SCREENSHOT)
				Set Jpeg = Nothing
			Else
				WEB_SITE_SCREENSHOT = "/none.jpg"
			End If
		Set objHTTP = Nothing
    End Property
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------


    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
	Public Function SSExist(vPath)
	    Dim Fso2 : Set Fso2 = Server.CreateObject("Scripting.FileSystemObject")
	    If Fso2.FileExists(Server.Mappath(vPath)) then
	        SSExist = True
	    Else
	        SSExist = False
	    End if
	    Set Fso2 = Nothing
	End Function
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------

    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
    Public Property Get GTMetrixIPListUpdate()
		If (DateDiff("d", CDate(GTMetrixIPLastFetch), Now()) < IPLIST_UPDATE_INTERVAL) Then
			Exit Property
		End If

	    Dim objXMLhttp : Set objXMLhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0") 
	        objXMLhttp.open "GET", GTMETRIX_FETCH_IP_URL, false
	        objXMLhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
	        objXMLhttp.setTimeouts 5000, 5000, 10000, 10000 'ms
	        objXMLhttp.setRequestHeader "user-agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36"
	        objXMLhttp.send

	        FindIPAddress objXMLhttp.responseText

	        UpdateSettings "GTMETRIX_BLOCKER_IPLIST", Join(GTMetrixIPCollection, ",")
	        UpdateSettings "GTMETRIX_BLOCKER_IPLIST_FETCHDATE", Cstr(Now())
	    Set objXMLhttp = Nothing
    End Property
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------

    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
    Public Property Get Check()
    	Dim VisitorIP
    	' For debug 
    	If Request.ServerVariables("http_gtmetrixblocker") = "test" OR ShowPreviewPage = True Then 
	    	HeaderExit 	= True
	    	SessionExit = False
	    	BotFound 	= True
	    	GTMetrixBot = True

	    	ScreenShot()
	    	Exit Property
    	End If

    	If TypeName(GTMetrixIP) = "Empty" Then 
    		Exit Property
    	End If

    	' Check Cache
    	If Session("GTMETRIX_BOT") = "-1" Then 
	    	SessionExit = True
	    	BotFound 	= False
	    	GTMetrixBot = False
	    	Exit Property
    	End If

		' Check UA
		If Instr(1, Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "GTmetrix") > 1 Then
	    	BotFound 	= True
	    	GTMetrixBot = True
	    	Exit Property
		End If

		' Check IP
	    VisitorIP 	= CheckIP()
	    GTMetrixBot = False

	    If in_array(VisitorIP, GTMetrixIP, True) = True Then
    		ScreenShot()
    		BotFound 	= True
	        GTMetrixBot = True
	        Exit Property 
	    End If
	    ' For i=0 To Ubound(GTMetrixIP)
		   '  If Trim(GTMetrixIP(i)) = Trim(VisitorIP) Then
	    ' 		ScreenShot()
	    		
	    ' 		BotFound 	= True
		   '      GTMetrixBot = True
		        
		   '      Exit Property      
		   '  End If    
	    ' Next
    End Property
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------


    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
	Private Function CheckIP()
		Dim t
	    	t = Request.ServerVariables("HTTP_CF-Connecting-IP") & ""
	    If Len(t) < 2 Then t = Request.ServerVariables("remote_addr")
	    CheckIP = t
	End Function
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------

    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
    Public Property Get BotPageHTML()
    	If BotFound = False Then 
    		Exit Property
    	End If

    	With Response 
			.Write "<!DOCTYPE html>"
			.Write "<html>"
			.Write "<head>"
			.Write "	<title>"& WEB_SITE_TITLE &"</title>"
			.Write "	<meta charset=""utf-8"">"
			.Write "	<style>header{display:block;text-align:center;background-color:#261e12;padding:10px;opacity:0.1}body{background-color:#f4f4f4;height:100%;position:relative} body>div{margin:auto;height:100vh;width:100vw;position:fixed;top:0;bottom:0;left:0;right:0;font-weight:bold;font-size:30px;background-image:url('"& WEB_SITE_SCREENSHOT &"');background-position:center top;background-repeat:no-repeat;background-size:cover;}.simplynone{font-size:12px;opacity:0.1; padding:200px;position:relative;width:100vw;height:100vh;}</style>"
			.Write "	<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css"" integrity=""sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm"" crossorigin=""anonymous"">"
			.Write "</head>"

    		.Write "<body class=""fully-load"">"
    		.Write "<div style=""min-height:400px;"">"
    		.Write "<header><a href=""/"" title=""home""><img alt="""& WEB_SITE_TITLE &""" width=""56"" height=""70"" src="""& WEB_SITE_LOGO &""" loading=""lazy"" /></a></header>"
    		.Write "<div class=""container""><div class=""row""><div class=""col-lg-12"">"
    		.Write "<div class=""simplynone"">"
    		.Write "<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur accumsan hendrerit risus non rutrum. Praesent ornare purus id risus luctus efficitur.</p> <p>Etiam sed eleifend nibh. Pellentesque sollicitudin risus ac libero tincidunt, vel egestas magna aliquet. Nullam scelerisque risus ut dolor placerat, eu lacinia mauris pharetra. Ut faucibus, massa sed volutpat pretium, augue enim accumsan massa, a porta erat nunc sed nibh.</p> <p>Duis id purus quis arcu rhoncus tempus. Donec non est eget neque tempor dictum. Phasellus eget erat arcu. Nullam consequat iaculis sapien, nec feugiat ex pellentesque ut. Donec rhoncus justo nibh, sed interdum lectus tristique in. Fusce porta justo non ex ullamcorper, et venenatis sapien rhoncus. Nam id vehicula risus, ut cursus tortor. Proin elementum massa odio, et bibendum dolor cursus at. Etiam sodales libero massa, id euismod erat scelerisque id.</p>"
    		.Write "</div>"
    		.Write "</div></div></div>"
    		.Write "</div>"
    		.Write "<footer> Copyright &copy; "& WEB_SITE_TITLE &" </footer>"
    		.Write "<script defer src=""https://code.jquery.com/jquery-3.2.1.slim.min.js"" integrity=""sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN"" crossorigin=""anonymous""></script>"
    		.Write "<script defer src=""https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/js/bootstrap.min.js"" integrity=""sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl"" crossorigin=""anonymous""></script>"
    		.Write "</body></html>"
    	End With
    	Response.End
	End Property
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------


    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
	Private Function FindIPAddress(Val)
	    Dim Text 
	        Text = Trim(Val) & ""

	    Text = Replace(Text, vbcrlf, "",1,-1,1)
	    Text = Replace(Text, vbcr, "",1,-1,1)
	    Text = Replace(Text, vblf, "",1,-1,1)
	    Text = Replace(Text, vbTab, "",1,-1,1)
	    Text = Replace(Text, "         ", "",1,-1,1)
	    Text = Replace(Text, "        ", "",1,-1,1)
	    Text = Replace(Text, "      ", "",1,-1,1)
	    Text = Replace(Text, "    ", "",1,-1,1)    
	    Text = Replace(Text, "   ", "",1,-1,1)    
	    Text = Replace(Text, "\n", "",1,-1,1)
	    Text = Replace(Text, "\r", "",1,-1,1)
	    Text = Replace(Text, """: """, """:""",1,-1,1)
	    Text = Replace(Text, """ />", """>",1,-1,1)

	    Dim objRegExp : Set objRegExp = New Regexp
	    With objRegExp
	        .Pattern = "<!--(?!<!)[^\[>].*?-->"
	        .IgnoreCase = False
	        .Global = True
	    End With
	    
	    Dim HTMLCompressor
	    	HTMLCompressor = objRegExp.Replace(ver,"")
	    
	    Set objRegExp = Nothing

	    Text = Replace(Text, vbcrlf, "")
	    Text = Replace(Text, vbcrlf, "")

		If ShowHTMLResult = True Then 
			Response.Write "<table class=""table table-striped"">"
		End If

		Dim TotalIP
			TotalIP = 0
	    Do While InStr(Text, "<label>IP:</label> ") > 0  AND InStr(Text, "<br> <label>Host:") > 0
	        DeyimBaslangici = InStr(Text, "<label>IP:</label>")
	        DeyimSonu = InStr(DeyimBaslangici, Text, "<br> <label>Host:") + 17
	        If DeyimSonu < DeyimBaslangici Then DeyimSonu = DeyimBaslangici + 17
	        strLink = Trim(Mid(Text, DeyimBaslangici, (DeyimSonu - DeyimBaslangici)))
	        strGeciciMesaj = strLink
	        strGeciciMesaj = Replace(strGeciciMesaj, "<label>IP:</label> ", "", 1, -1, 1)
	        strGeciciMesaj = Replace(strGeciciMesaj, "<br> <label>Host:", "", 1, -1, 1)
	        strGeciciMesaj = Trim(strGeciciMesaj)
	        If Isnull(strGeciciMesaj) Or IsEmpty(strGeciciMesaj) OR strGeciciMesaj = "" Then
	            strGeciciMesaj = "#"
	        Else
	        	TotalIP = TotalIP + 1
	            AddToArray GTMetrixIPCollection, Trim(strGeciciMesaj)
	            strGeciciMesaj = "<tr> <td>"& strGeciciMesaj &"</td> <td>IP Adress Saved<td> <tr>"
	            If ShowHTMLResult = True Then 
	            	Response.Write strGeciciMesaj
	            End If
	        End If
	        Text = Replace(Text, strLink, strGeciciMesaj, 1, -1, 1)
	    Loop

	    If ShowHTMLResult = True Then 
	    	If TotalIP = 0 Then
				Response.Write "<tr><td colspan=""2"" align=""center""><strong style=""color:red"">Hiç Bir Adres Alınmadı. Veri sayfasını kontrol edin. "& GTMETRIX_FETCH_IP_URL &"</strong></td></tr>" & vbcrlf
	    	Else
				Response.Write "<tr><td colspan=""2"" align=""center""><strong>Bütün Adreslen Alındı! Lütfen Eklenti Panelini Yeniden Yükleyin!</strong></td></tr>" & vbcrlf
			End If
			Response.Write "</table>"
		End If
	    FindIPAddress = Text
	End Function
    '------------------------------------------------------------------------------------------
    ' 
    '------------------------------------------------------------------------------------------
End Class
%>
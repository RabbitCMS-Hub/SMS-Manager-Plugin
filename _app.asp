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

Dim SMSWords()
    SMSWordsSize=0

Const SMS_ERROR 	= 0
Const SMS_SUCCESS 	= 1
Const PACKET_XML 	= 0
Const PACKET_JSON 	= 1

Class SMS_Manager_Plugin
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME, PLUGIN_AUTOLOAD

	Private SYSTEM_LANG_ID
	Private SMS_MODULE_ACTIVE, SMS_MODULE_APIKEY, SMS_MODULE_HEADER, SMS_MODULE_PROVIDER_URL, SMS_MODULE_PROVIDER, SMS_MODULE_PASSWORD
	Private SMS_STATUS, SMS_RESPONSE
	Private SMS_MESAJ, SMS_ALICI, BYPASS_MODULE, SEND_REQUEST, GET_RESPONSE, SMS_TEST_MESSAGE_TEXT

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

		' Check And Create Table
		'------------------------------
		Dim PluginTableName
			PluginTableName = "tbl_plugin_" & PLUGIN_DB_NAME


		' Check And Create Table
		'------------------------------
    	If TableExist(PluginTableName) = False Then
    		Conn.Execute("SET NAMES utf8mb4;") 
    		Conn.Execute("SET FOREIGN_KEY_CHECKS = 0;") 
    		
    		Conn.Execute("DROP TABLE IF EXISTS `"& PluginTableName &"_module`")
    		q="CREATE TABLE `"& PluginTableName &"_module` ( "
    		q=q+"  `ID` int(11) NOT NULL AUTO_INCREMENT, "
    		q=q+"  `NUMARA` varchar(20) DEFAULT NULL, "
    		q=q+"  `MESAJ` varchar(255) DEFAULT NULL, "
    		q=q+"  `GONDERIM_TARIHI` datetime DEFAULT NULL, "
    		q=q+"  `PROVIDER` varchar(20) DEFAULT NULL, "
    		q=q+"  PRIMARY KEY (`ID`), "
    		q=q+"  KEY `IND1` (`NUMARA`) "
    		q=q+") ENGINE=MyISAM DEFAULT CHARSET=utf8; "
			Conn.Execute(q) : q=""
    		Call PanelLog(""& PLUGIN_CODE &"_module için database tablosu oluşturuldu", 0, ""& PLUGIN_CODE &"", 0)

    		Conn.Execute("DROP TABLE IF EXISTS `"& PluginTableName &"_block`")
    		q="CREATE TABLE `"& PluginTableName &"_block` ( "
    		q=q+"  `ID` int(11) NOT NULL AUTO_INCREMENT, "
    		q=q+"  `NUMARA` varchar(20) DEFAULT NULL, "
    		q=q+"  `TARIH` datetime DEFAULT NULL, "
    		q=q+"  PRIMARY KEY (`ID`), "
    		q=q+"  KEY `IND1` (`NUMARA`) "
    		q=q+") ENGINE=MyISAM DEFAULT CHARSET=utf8; "
			Conn.Execute(q) : q=""

    		Conn.Execute("SET FOREIGN_KEY_CHECKS = 1;") 
    		Call PanelLog(""& PLUGIN_CODE &"_block için database tablosu oluşturuldu", 0, ""& PLUGIN_CODE &"", 0)

    	End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE)
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "SMS_Manager_Plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "589")
		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "0")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)


		' Plugin Settings
		'------------------------------
		a=GetSettings(""&PLUGIN_CODE&"_PROVIDER", "SANALSANTRAL")
		a=GetSettings(""&PLUGIN_CODE&"_PROVIDER_URL", "")
		a=GetSettings(""&PLUGIN_CODE&"_APIKEY", "")
		a=GetSettings(""&PLUGIN_CODE&"_PASSWORD", "")
		a=GetSettings(""&PLUGIN_CODE&"_HEADER", "")

		a=GetSettings(""&PLUGIN_CODE&"_PREMSG_USERREGISTER", "Hoşgeldiniz [ADI]. Kayıt işleminiz başarılı bir şekilde tamamlanmıştır.")
		a=GetSettings(""&PLUGIN_CODE&"_PREMSG_USERREGISTER_STATUS", "0")
		a=GetSettings(""&PLUGIN_CODE&"_PREMSG_LOSTPASSWD", "Yeni parolanız [YENIPAROLA] olarak tanımlanmıştır. Parolanız büyük küçük harf duyarlıdır.")
		a=GetSettings(""&PLUGIN_CODE&"_PREMSG_LOSTPASSWD_STATUS", "0")
		a=GetSettings(""&PLUGIN_CODE&"_PREMSG_ORDERCOMPLETE", "[SIPARISNO] numaralı siparişiniz alındı. Teşekkür ederiz.")
		a=GetSettings(""&PLUGIN_CODE&"_PREMSG_ORDERCOMPLETE_STATUS", "0")

		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Plugin Admin Panel Extention
	'---------------------------------------------------------------
	Public sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "SHOW:BlockedNumbers" Then
			Call PluginPage("Header")

			With Response
				.Write "<div class=""table-responsive"">"
				.Write "    <table class=""table table-striped table-bordered"">"
				.Write "        <thead>"
				.Write "            <tr>"
				.Write "                <th>Numara</th>"
				.Write "                <th>İşlem Tarihi</th>"
				.Write "                <th></th>"
				.Write "            </tr>"
				.Write "        </thead>"
				.Write "        <tbody>"
				Set Siteler = Conn.Execute("SELECT * FROM tbl_plugin_sms_block ORDER BY ID DESC")
				If Siteler.Eof Then 
				    Response.Write "<tr>"
				        Response.Write "<td colspan=""3"" align=""center"">İşlem Geçmişi Bulunamadı</td>"
				    Response.Write "</tr>"
				End If
				Do While Not Siteler.Eof
				.Write "            <tr>"
				.Write "                <td>"& Siteler("NUMARA") &"</td>"
				.Write "                <td>"& Siteler("TARIH") &"</td>"
				.Write "                <td></td>"
				.Write "            </tr>"
				Siteler.MoveNext : Loop
				Siteler.Close : Set Siteler = Nothing
				.Write "        </tbody>"
				.Write "    </table>"
				.Write "</div>"
			End With

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "SHOW:SMSLog" Then
			Call PluginPage("Header")

			With Response
				.Write "<div class=""table-responsive"">"
				.Write "    <table class=""table table-striped table-bordered"">"
				.Write "        <thead>"
				.Write "            <tr>"
				.Write "                <th>Numara</th>"
				.Write "                <th>Mesaj</th>"
				.Write "                <th>Tarih</th>"
				.Write "                <th>Servis Sağlayıcı</th>"
				.Write "            </tr>"
				.Write "        </thead>"
				.Write "        <tbody>"
				Set Siteler = Conn.Execute("SELECT * FROM tbl_plugin_sms_module ORDER BY ID DESC")
				If Siteler.Eof Then 
				    Response.Write "<tr>"
				        Response.Write "<td colspan=""4"" align=""center"">İşlem Geçmişi Bulunamadı</td>"
				    Response.Write "</tr>"
				End If
				Do While Not Siteler.Eof
				.Write "            <tr>"
				.Write "                <td>"& Siteler("NUMARA") &"</td>"
				.Write "                <td>"& Siteler("MESAJ") &"</td>"
				.Write "                <td>"& Siteler("GONDERIM_TARIHI") &"</td>"
				.Write "                <td>"& Siteler("PROVIDER") &"</td>"
				.Write "            </tr>"
				Siteler.MoveNext : Loop
				Siteler.Close : Set Siteler = Nothing
				.Write "        </tbody>"
				.Write "    </table>"
				.Write "</div>"
			End With

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "AJAX:TestSMS" Then
		    Query.PageContentType = "json"
			
			Set rsTempClass = New SMS_Manager_Plugin
				rsTempClass.ByBass = 1
				rsTempClass.Mesaj  = SMS_TEST_MESSAGE_TEXT
				rsTempClass.Numara = Query.Data("NUMARA")
				
				apiResponse = rsTempClass.SendSMS()
					tmp_status 	= apiResponse(0)
					tmp_response= apiResponse(1)

				If tmp_status = SMS_SUCCESS Then 
		        	Query.jsonResponse 200, tmp_response
				Else 
		        	Query.jsonResponse 400, tmp_response
				End If
			Set rsTempClass = Nothing

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
			.Write 			QuickSettings("select", ""& PLUGIN_CODE &"_PROVIDER", "Servis Sağlayıcı", "SANALSANTRAL#Sanal Santral|CAGRISMS#Çağrı SMS|NETGSM#NetGSM", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", ""& PLUGIN_CODE &"_PROVIDER_URL", "Servis Sağlayıcı URL (Boş Default Kullanır)", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("input", ""& PLUGIN_CODE &"_APIKEY", "API Anahtarı veya Kullanıcı Adı", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("input", ""& PLUGIN_CODE &"_PASSWORD", "API Şifresi (opsiyonel)", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("input", ""& PLUGIN_CODE &"_HEADER", "SMS Header", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("textarea", ""& PLUGIN_CODE &"_PREMSG_USERREGISTER", "Yeni Üyeye Giden Mesaj", "", TO_DB)
			.Write 			QuickSettings("checkbox", ""& PLUGIN_CODE &"_PREMSG_USERREGISTER_STATUS", "Aktiflik", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("textarea", ""& PLUGIN_CODE &"_PREMSG_LOSTPASSWD", "Parolamı Unuttum Giden Mesaj", "", TO_DB)
			.Write 			QuickSettings("checkbox", ""& PLUGIN_CODE &"_PREMSG_LOSTPASSWD_STATUS", "Aktiflik", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-4 col-sm-12"">"
			.Write 			QuickSettings("textarea", ""& PLUGIN_CODE &"_PREMSG_ORDERCOMPLETE", "Sipariş Tamamlandı Giden Mesaj", "", TO_DB)
			.Write 			QuickSettings("checkbox", ""& PLUGIN_CODE &"_PREMSG_ORDERCOMPLETE_STATUS", "Aktiflik", "", TO_DB)
			.Write "    </div>"

			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write " 		<strong>Kullanılabilir Kısaltmalar</strong><code>[DOMAIN],[TARIH],[ADSOYAD],[YENIPAROLA],[SIPARISNO]</code>"
			.Write "    </div>"
			.Write "</div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12 mt-3"" style=""background-color:rgb(55 227 255 / 30%);padding:20px;margin:0px 15px 10px 15px;"">"
		    .Write "<form method=""post"" action=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=AJAX:TestSMS"" data-ajax-submit data-ajax-callback=""SMSTestResult"" data-ajax-modalclose=""false"">"
			.Write "    <div class=""form-group"">"
			.Write "    	<label>SMS Test Mesajı</label>"
			.Write "        <input type=""text"" class=""form-control"" value="""& SMS_TEST_MESSAGE_TEXT &""" readonly />"
			.Write "    </div>"
			.Write "    <div class=""form-group mb-0"">"
			.Write "    	<label>SMS Gönderim Testi</label>"
			.Write "        <div class=""input-group"">"
			.Write "    		<div class=""form-group"">"
			.Write "            	<input name=""NUMARA"" placeholder=""Örn. 5122345678"" type=""text"" minlength=""10"" maxlength=""10"" pattern=""^[5][0-9]{9}$"" class=""form-control"" required autocomplete=""off"" />"
			.Write "            </div>"
		    .Write "            <button type=""submit"" id=""SendSMSPreview"" onclick=""$('#SMSTestSonucu').html('İşlem yapılıyor...');"" class=""btn btn-success"">"
		    .Write "                Test Gönder"
		    .Write "            </button>"
			.Write "        </div>"
			.Write "        <div><small id=""SMSTestSonucu"">Test sonucu bekleniyor</small></div>"
			.Write "    </div>"
		    .Write "</form>"
			.Write "    </div>"
			.Write "</div>"

			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:BlockedNumbers"" class=""btn btn-sm btn-primary"">"
			.Write "        	Bloklanan Numaraları Görüntüle"
			.Write "        </a>"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:SMSLog"" class=""btn btn-sm btn-primary"">"
			.Write "        	Logları Görüntüle"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
			
	    	.Write "<script type=""text/javascript"" src=""/content/plugins/SMS-Manager-Plugin/js/app.js""></script>"
		End With
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Class First Init
	'---------------------------------------------------------------
	Private Sub class_initialize()
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_CODE  			= "SMS_MODULE"
    	PLUGIN_NAME 			= "SMS Manager Plugin"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/SMS-Manager-Plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_ICON 			= "zmdi-local-post-office"
    	PLUGIN_CREDITS 			= "@badursun Anthony Burak DURSUN"
    	PLUGIN_FOLDER_NAME 		= "SMS-Manager-Plugin"
    	PLUGIN_DB_NAME 			= "sms" ' tbl_plugin_XXXXXXX
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_AUTOLOAD 		= False
    	PLUGIN_ROOT 			= PLUGIN_DIST_FOLDER_PATH(This)
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------

        SYSTEM_LANG_ID          = Session("LANG_ID")
        If IsNull(SYSTEM_LANG_ID) OR IsEmpty(SYSTEM_LANG_ID) OR Not IsNumeric(SYSTEM_LANG_ID) OR SYSTEM_LANG_ID="" Then
            SYSTEM_LANG_ID = SYSTEM_DEFAULT_LANG_ID
        End If

		SMS_MODULE_ACTIVE 		= Cint( GetSettings("SMS_MODULE_ACTIVE", "0") )
		SMS_MODULE_APIKEY 		= GetSettings("SMS_MODULE_APIKEY", "")
		SMS_MODULE_PASSWORD 	= GetSettings("SMS_MODULE_PASSWORD", "")
		SMS_MODULE_HEADER 		= GetSettings("SMS_MODULE_HEADER", "")
		SMS_MODULE_PROVIDER 	= GetSettings("SMS_MODULE_PROVIDER", "SANALSANTRAL")
		SMS_MODULE_PROVIDER_URL = GetSettings("SMS_MODULE_PROVIDER_URL", "")
		SMS_STATUS 				= 0
		SMS_RESPONSE 			= Null
		SMS_MESAJ 				= Null
		SMS_ALICI 				= Null
		BYPASS_MODULE 			= 0
		SEND_REQUEST 			= Null
		GET_RESPONSE 			= Null
		SMS_TEST_MESSAGE_TEXT 	= ""& PLUGIN_NAME &" Test Messages Send @"& Now() &" From:"& IPAdresi() &" On:"& DOMAIN_URL &" ("& SETTINGS_CMS_UNIQUE_ID &")"

    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Register App
    	'-------------------------------------------------------------------------------------
    	class_register()

    	'-------------------------------------------------------------------------------------
    	' Hook Auto Load Plugin
    	'-------------------------------------------------------------------------------------
    	If PLUGIN_AUTOLOAD_AT("WEB") = True Then 

    	End If
	End Sub
	'---------------------------------------------------------------
	' Class First Init
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Class Terminate
	'---------------------------------------------------------------
	Private sub class_terminate()

	End Sub
	'---------------------------------------------------------------
	' Class Terminate
	'---------------------------------------------------------------


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


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let ByBass(vVal)
		BYPASS_MODULE = vVal
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get RequestData()
		RequestData = SEND_REQUEST
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get ResponseData()
		ResponseData = GET_RESPONSE
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get SendSMS()
		' Error Handling
		'------------------------------------------------
		If SMS_MODULE_ACTIVE = 0 AND BYPASS_MODULE = 0 Then 
			SendSMS = Array(SMS_ERROR, "SMS Modülü Pasif") : Exit Property
		End If
		If Len(SMS_MODULE_APIKEY)=0 Then 
			SendSMS = Array(SMS_ERROR, "Geçersiz API Anahtarı veya Kullanıcı Adı") : Exit Property
		End If
		If Len(SMS_MODULE_PASSWORD)=0 Then 
			SendSMS = Array(SMS_ERROR, "Geçersiz API Parolası") : Exit Property
		End If
		If Len(SMS_MODULE_HEADER)=0 Then 
			SendSMS = Array(SMS_ERROR, "Geçersiz SMS Header") : Exit Property
		End If

		If Len(SMS_MESAJ)<1 Then 
			SendSMS = Array(SMS_ERROR, "Boş SMS Gönderilmemeli") : Exit Property
		End If
		If Len(SMS_ALICI)<10 Then 
			SendSMS = Array(SMS_ERROR, "Geçersiz SMS Alıcı Numarası") : Exit Property
		End If

		' Send By Selected Provider
		'------------------------------------------------
		Select Case SMS_MODULE_PROVIDER
			Case "SANALSANTRAL" 	: SendSMS = SendFromTescom()
			Case "CAGRISMS" 		: SendSMS = SendFromCagriSMS()
			Case "NETGSM" 			: SendSMS = SendFromNetGsm()
			Case Else
				SendSMS = Array(SMS_ERROR, "Servis Sağlayıcı Bulunamadı ("& SMS_MODULE_PROVIDER &")")
		End Select
		

	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let Numara(vTxt)
		SMS_ALICI = FormatPhoneNumber( vTxt )
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let Mesaj(vTxt)
		SMS_MESAJ = ReplaceTurkishChar( SMSReplacer(vTxt) )
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function ReplaceTurkishChar(Txt)   
		Dim rtChar
			rtChar = Txt

		rtChar = Replace(rtChar,"ç" ,"c" )
		rtChar = Replace(rtChar,"ğ" ,"g" )
		rtChar = Replace(rtChar,"ü" ,"u" )
		rtChar = Replace(rtChar,"ş" ,"s" )
		rtChar = Replace(rtChar,"ö" ,"o" )
		rtChar = Replace(rtChar,"İ" ,"I" )
		rtChar = Replace(rtChar,"ı" ,"i" )
		rtChar = Replace(rtChar,"Ğ" ,"G" )
		rtChar = Replace(rtChar,"Ü" ,"U" )
		rtChar = Replace(rtChar,"Ş" ,"S" )
		rtChar = Replace(rtChar,"Ö" ,"O" )
		rtChar = Replace(rtChar,"Ç" ,"C" )
		ReplaceTurkishChar = rtChar
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function Cagri90(No)
		Dim PhoneNo 
			PhoneNo = No & "" 
		If Left(PhoneNo,2) = "90" Then 
			Cagri90 = Mid(PhoneNo, 3, Len(PhoneNo))
		Else 
			Cagri90 = PhoneNo
		End If
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function FormatPhoneNumber(vTxt)    
		Dim FPNumber
			FPNumber = vTxt
		
		If Instr(1, FPNumber, ",") <> 0 Then 
			tmp_split = Split(FPNumber, ",")
			For y=0 To UBound(tmp_split)
				tmp_split(y) = FixPhoneNumber( tmp_split(y) )
			Next
			FPNumber = Join(tmp_split, ",")
		Else 
			FPNumber = FixPhoneNumber(FPNumber)
		End If

		FormatPhoneNumber = Trim(FPNumber)
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
    Public Function FixPhoneNumber(vVal)
		Dim FormattedNumber
		vVal = Replace(vVal, " ", "")
		
		If Len(vVal) = 11 AND Left(vVal,1) = "0" Then
			FormattedNumber = Right(vVal, Len(vVal)-1 )
		Else
			FormattedNumber = vVal
		End If

		If IsBlocked(FormattedNumber)=True Then
			FixPhoneNumber = "0"
			Exit Function
		End If

		FixPhoneNumber = FormattedNumber
    End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Property Get XMLHttp(Uri, xType, Data, PacketType)
		' Data To Variable
		'------------------------------------------------
		SEND_REQUEST = Data

		' ByPass Error
		'------------------------------------------------
		On Error Resume Next

		' Override Default Provider URL
		'------------------------------------------------
		If Len(SMS_MODULE_PROVIDER_URL) > 10 Then 
			Uri = SMS_MODULE_PROVIDER_URL
		End If

		' Send Data
		'------------------------------------------------
	    Set objXMLhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
            objXMLhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
            objXMLhttp.setTimeouts 5000, 5000, 10000, 10000 'ms
            objXMLhttp.setRequestHeader "X-Cms-UniqueId", SETTINGS_CMS_UNIQUE_ID
			objXMLhttp.open xType, Uri, false

			Select Case PacketType
				Case PACKET_JSON
					objXMLhttp.setRequestHeader "Accept","application/json"
					objXMLhttp.setRequestHeader "Content-type","application/json"
				Case PACKET_XML
					objXMLhttp.setRequestHeader "charset","utf-8"
					objXMLhttp.setRequestHeader "Content-type","application/xml"
				Case Else

			End Select

			objXMLhttp.send Data
			
			GET_RESPONSE = objXMLhttp.responseText
			XMLHttp = Array(objXMLhttp.Status, objXMLhttp.responseText)
	    
	    Set objXMLhttp = Nothing
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function SendFromNetGsm()
		Dim URL_NETGSM
			URL_NETGSM = "https://api.netgsm.com.tr/sms/send/xml"

	    xml="<?xml version=""1.0"" encoding=""UTF-8""?>"
	    xml=xml+"<mainbody> "
		    xml=xml+"<header> "
			    xml=xml+"<company dil=""TR"">Netgsm</company>  "
			    xml=xml+"<usercode>"& SMS_MODULE_APIKEY &"</usercode> "
			    xml=xml+"<password>"& SMS_MODULE_PASSWORD &"</password> "
			    xml=xml+"<type>1:n</type> "
			    xml=xml+"<msgheader>"& SMS_MODULE_HEADER &"</msgheader> "
		    xml=xml+"</header> "
		    xml=xml+"<body> "
	    	
	    	xml=xml+"<msg><![CDATA["& SMS_MESAJ &"]]></msg> "

		    If Instr(1, SMS_ALICI, ",") <> 0 Then 
		    	tmp_numaralar = Split(SMS_ALICI, ",")
		    	For smsNo=0 To UBound(tmp_numaralar)
		    		xml=xml+"<no>"& tmp_numaralar(smsNo) &"</no> "

					' Log
					'--------------------------------------------
					SMSLog tmp_numaralar(smsNo), SMS_MESAJ
		    	Next
		    Else
		    	xml=xml+"<no>"& SMS_ALICI &"</no> "

				' Log
				'--------------------------------------------
				SMSLog SMS_ALICI, SMS_MESAJ
			End If
		    
		    xml=xml+"</body> "
		xml=xml+"</mainbody> "

		Response.Write xml 
	    SendFromNetGsm = XMLHttp(URL_NETGSM, "POST", xml, PACKET_XML) : xml=""
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function SendFromTescom()
		Dim URL_TESCOM
			URL_TESCOM = "http://sms2.sanalsantral.com.tr/api/smspost/v1"

	    xml=""
	    xml=xml+"<sms> "
		    xml=xml+"<header>"& SMS_MODULE_HEADER &"</header>"
		    xml=xml+"<apikey>"& SMS_MODULE_APIKEY &"</apikey>"
		    xml=xml+"<validity>2880</validity>"
		    xml=xml+"<type>1</type>"
		    xml=xml+"<message>"
			    xml=xml+"<gsm>"
			    If Instr(1, SMS_ALICI, ",") <> 0 Then 
			    	tmp_numaralar = Split(SMS_ALICI, ",")
			    	For smsNo=0 To UBound(tmp_numaralar)
			    		xml=xml+"<no>"& tmp_numaralar(smsNo) &"</no>"

    					' Log
    					'--------------------------------------------
    					SMSLog tmp_numaralar(smsNo), SMS_MESAJ
			    	Next
			    Else
			    	xml=xml+"<no>"& SMS_ALICI &"</no>"

					' Log
					'--------------------------------------------
    				SMSLog SMS_ALICI, SMS_MESAJ
				End If
			    xml=xml+"</gsm>"
		    	xml=xml+"<msg><![CDATA["& SMS_MESAJ &"]]></msg>"
		    xml=xml+"</message>"
	    xml=xml+"</sms>"

	    SendFromTescom = XMLHttp(URL_TESCOM, "POST", xml, PACKET_XML) : xml=""
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Function SendFromCagriSMS()
		Dim URL_CAGRISMS
			URL_CAGRISMS = "https://api.cagrisms.com/api/SMS"

		Set oJSON = New aspJSON
		With oJSON.data
			.Add "header", oJSON.Collection()
		    With oJSON.data("header")
				.Add "username", ""& SMS_MODULE_APIKEY &""
				.Add "password", ""& SMS_MODULE_PASSWORD &""
		    End With

			.Add "body", oJSON.Collection()
		    With oJSON.data("body")
			    .Add "messages", oJSON.Collection()
			    With .item("messages")
			    	If Instr(1, SMS_ALICI, ",") <> 0 Then 
				    	tmp_numaralar = Split(SMS_ALICI, ",")
				    	For smsNo=0 To UBound(tmp_numaralar)
					    	.Add smsNo, oJSON.Collection()
					    	With .item(smsNo)
								.Add "phoneNumber", ""& Cagri90( tmp_numaralar(smsNo) ) &""
								.Add "content", ""& SMS_MESAJ &""
					    	End With

	    					' Log
	    					'--------------------------------------------
    						SMSLog tmp_numaralar(smsNo), SMS_MESAJ
				    	Next
			    	Else
				    	.Add 0, oJSON.Collection()
				    	With .item(0)
							.Add "phoneNumber", ""& Cagri90( SMS_ALICI ) &""
							.Add "content", ""& SMS_MESAJ &""
				    	End With
    					
    					' Log
    					'--------------------------------------------
    					SMSLog SMS_ALICI, SMS_MESAJ
					End If
			    End With
				.Add "title", ""& SMS_MODULE_HEADER &""
		    End With
		End With

		' Return Response JSON
		'----------------------------------------------
		Dim smsData
			smsData = oJSON.JSONoutput()

		' Destroy Json
		'----------------------------------------------
		Set oJSON = Nothing

		' İşlem sonucu
		'----------------------------------------------
	    tmp_response = XMLHttp(URL_CAGRISMS, "POST", smsData, PACKET_JSON) : smsData=""
		
		' 200 Harici Sunucu kodu döndüyse
		'----------------------------------------------
	    If Not tmp_response(0) = "200" Then 
			SendFromCagriSMS = Array(SMS_ERROR, "Bir hata oluştu. Yanıt kodu:"& tmp_response(0) &" Mesaj:"& tmp_response(1) &"" )
			Exit Function
	    End If

		' Yanıtı parçala ve ayrıştır
		'----------------------------------------------
		Set parseJsonData = New aspJSON
			parseJsonData.loadJSON( tmp_response(1) )
		
			If IsNull(parseJsonData.data("success")) = True Then 
				SendFromCagriSMS = Array(SMS_ERROR, parseJsonData.data("message") )
			Else 
				SendFromCagriSMS = Array(SMS_SUCCESS, "İşlem başarılı. Paket kodu: " + parseJsonData.data("data").item("messagePackageGuid") )
			End If
		Set parseJsonData = Nothing

	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Property Get SMSLog(No, Msg)
    	Conn.Execute("INSERT INTO tbl_plugin_sms_module(NUMARA, MESAJ, GONDERIM_TARIHI, PROVIDER) VALUES('"& No &"', '"& Msg &"', NOW(), '"& SMS_MODULE_PROVIDER &"')")
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get Replacer(Anahtar, Yerine)
	    If IsNull(Anahtar) AND IsNull(Yerine) Then 
	        Erase SMSWords
	    Else
	        ReDim PRESERVE SMSWords(HerEklenenVeriSize)
	        SMSWordsSize=SMSWordsSize+1

	        ReDim PRESERVE SMSWords( HerEklenenVeriSize )

	        SMSWords(SMSWordsSize-1) = ""&Anahtar&"!|!"&Yerine&""
	    End If
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function SMSReplacer(gelenveri) 
	    gelenveri = Replace(gelenveri, "[DOMAIN]", 							GetDomain(),1,-1,1)
	    gelenveri = Replace(gelenveri, "[TARIH]", 							Now(),1,-1,1)

	    If SMSWordsSize > 0 Then 
	        For EWS=0 To UBound( SMSWords )
	            HerEklenenVeri = Split( SMSWords(EWS), "!|!")
	            If UBound(HerEklenenVeri) = 1 Then 
	                gelenveri = Replace(gelenveri, Trim( HerEklenenVeri(0) ), Trim( HerEklenenVeri(1) ),1,-1,1)
	            End If
	        Next
	    End If
	    SMSReplacer = HTMLtoOneLine(gelenveri)
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function IsBlocked(Number)
		IsBlocked = True
		Set rsCheckNumber = Conn.Execute("SELECT ID FROM tbl_plugin_sms_block WHERE NUMARA='"& Number &"'")
			If rsCheckNumber.Eof Then IsBlocked = False
		rsCheckNumber.Close : Set rsCheckNumber = Nothing
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
End Class 
%>
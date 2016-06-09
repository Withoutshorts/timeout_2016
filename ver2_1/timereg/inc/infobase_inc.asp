<%

	sub deveditset
	'************************************* DEVEDIT SETTINGS ************************************
	
			'Create a new DevEdit class object
			dim myDE
			set myDE = new DevEdit
			
			'Set de path
			SetDevEditPath("de")
			
			'Set the name of this DevEdit class
			myDE.SetName "myDevEditTimeout"
			
			
			'Set the path to the "de" folder
			'The path can be a relative or absolute path. A relative path however CANNOT contain '../'
			'eg. Absolute path -> '/mycms/devedit/de'
			'    Relative path -> 'mycms/de'
			'    Relative path -> 'de'
			SetDeveditPath "de"
		
			'Set the path to the folder that contains the flash files for the flash manager
			myDE.SetFlashPath("/devedit/flash_test")
		
			'These are the functions that you can call to hide varions buttons,
			'lists and tab buttons. By default, everything is enabled
		
			myDE.HideSaveButton
			myDE.HideFullScreenButton
			'myDE.HideBoldButton
			'myDE.HideUnderlineButton
			'myDE.HideItalicButton
			'myDE.HideStrikethroughButton
			'myDE.HideNumberListButton
			'myDE.HideBulletListButton
			'myDE.HideDecreaseIndentButton
			'myDE.HideIncreaseIndentButton
			'myDE.HideLeftAlignButton
			'myDE.HideCenterAlignButton
			'myDE.HideRightAlignButton
			'myDE.HideJustifyButton
			'myDE.HideHorizontalRuleButton
			'myDE.HideLinkButton
			myDE.HideAnchorButton
			'myDE.HideMailLinkButton
			'myDE.HideHelpButton
			'myDE.HideFontList
			'myDE.HideSizeList
			'myDE.HideFormatList
			myDE.HideStyleList
			'myDE.HideForeColorButton
			'myDE.HideBackColorButton
			'myDE.HideTableButton
			myDE.HideFormButton
			'myDE.HideImageButton
			myDE.HideFlashButton
			myDE.DisableFlashUploading
			myDE.DisableFlashDeleting
			myDE.DisableInsertFlashFromWeb
			myDE.HideTextBoxButton
			'myDE.HideSymbolButton
			'myDE.HidePropertiesButton
			myDE.HideCleanHTMLButton
			myDE.HideAbsolutePositionButton
			'myDE.HideGuidelinesButton
			'myDE.HideSpellingButton
			'myDE.HideRemoveTextFormattingButton
			'myDE.HideSuperScriptButton
			'myDE.HideSubScriptButton
			'myDE.DisableSourceMode
			'myDE.DisablePreviewMode
			'myDE.DisableImageUploading
			'myDE.DisableImageDeleting
			'myDE.DisableInsertImageFromWeb
			'myDE.DisableXHTMLFormatting
			'myDE.DisableSingleLineReturn
			
			'If you want to use the spell checker, then you can set
			'the spelling language to de_AMERICAN, de_BRITISH or de_CANADIAN,
			'de_FRENCH, de_SPANISH, de_GERMAN, de_ITALIAN, de_PORTUGESE,
			'de_DUTCH, de_NORWEGIAN, de_SWEDISH or de_DANISH
			myDE.SetLanguage de_DANISH
		
			'We can specify a list of fonts for the font drop down. If we don't,
			'then a default list will show
			'myDE.SetFontList ("Verdana, Arial, Courier")
		
			'We can specify a list of font sizes for the font size drop down. If we don't,
			'then a default list will show
			'myDE.SetFontSizeList "2,3,4,5"
		
			'How do we want images to be inserted into our HTML content?
			'de_PATH_TYPE_FULL will insert a link/image in this format: http://www.mysite.com/test.html
			'de_PATH_TYPE_ABSOLUTE will insert a link/image in this format: /myimage.gif
			myDE.SetPathType de_PATH_TYPE_ABSOLUTE
			'myDE.SetPathType de_PATH_TYPE_FULL
			
	
			'Are we editing a full HTML page, or just a snippet of HTML?
			'de_DOC_TYPE_HTML_PAGE means we're editing a complete HTML page
			'de_DOC_TYPE_SNIPPET means we're editing a snippet of HTML
			myDE.SetDocumentType de_DOC_TYPE_SNIPPET
			'myDE.SetDocumentType de_DOC_TYPE_HTML_PAGE
		
		
			'Do we want images to appear in the image manager as thumbnails or just in rows?
			'de_IMAGE_TYPE_ROW means just list in a tabular format, without a thumbnail
			'de_IMAGE_TYPE_THUMBNAIL means list in 4-per-line thumbnail mode
			'myDE.SetImageDisplayType de_IMAGE_TYPE_ROW
			myDE.SetImageDisplayType de_IMAGE_TYPE_THUMBNAIL
		
			'How do we want flash movies to appear in the flash manager? As thumbnails or just in rows?
			'de_FLASH_TYPE_ROW means just list in a tabular format, without a thumbnail
			'de_FLASH_TYPE_THUMBNAIL means list in a 4-per-line thumbnail mode
			'myDE.SetFlashDisplayType de_FLASH_TYPE_ROW
			myDE.SetFlashDisplayType de_FLASH_TYPE_THUMBNAIL
		
			'Show table guidelines as dashed
			myDE.EnableGuidelines
			
			'If the user isnt running Internet Explorer, then a <textarea> tag will be shown.
			'This function will set the rows and cols of that <textarea>
			myDE.SetTextAreaDimensions 60, 30
		
			'If in snippet mode, specify a base href so that images are displayed
			'myDE.SetBaseHref "http://www.devedit.com"
			
			dim val
			
			if myDE.GetValue(false) = "" then
				val = strBesk
			else
				val = myDE.GetValue(false)
			end if
			
			'Set the initial HTML value of our control
			myDE.SetValue ("<font face='Arial' size='2'>"&val)
			
			
			' Use the LoadHTMLFromAccessQuery function to load a value based on a query
			' dim errDesc
			' errDesc = ""
			' myDE.LoadHTMLFromAccessQuery "c:\test.mdb", "select testField from testTable", errDesc
			' if errDesc <> "" then
			'	  Response.Write "An error occured: " & errDesc
			' end if
			
			' Use the LoadHTMLFromSQLServerQuery function to load a value based on a query
			' dim errDesc
			' errDesc = ""
			' myDE.LoadHTMLFromSQLServerQuery "localhost", "testDatabase", "testUser", "testPassword", "select testField from testTable", errDesc
			' if errDesc <> "" then
			'	  Response.Write "An error occured: " & errDesc
			' end if
			
			' Use the LoadFromFile function to load a complete text file
			'dim errDesc
			'errDesc = ""
			'myDE.LoadFromFile "timereg.asp", errDesc
			'if errDesc <> "" then
				 'Response.Write "An error occured: " & errDesc
			'end if
			
			' dim errDesc
			' errDesc = ""
			' myDE.SaveToFile "c:\testfile.txt", errDesc
			' if errDesc <> "" then
			'   Response.Write "An error occured: " & errDesc
			' end if
			
			' Use the AddCustomInsert function to add some custom inserts
			'myDE.AddCustomInsert "DevEdit Logo", "<img src='http://www.devedit.com/images/logo.gif'>"
			'myDE.AddCustomInsert "Red Text", "<font face='verdana' color='red' size='3'><b>Red Text</b></font>"
		
			' Use the AddCustomLink function to add some custom links
			'myDE.AddCustomLink "DevEdit Website", "http://www.devedit.com", "_blank"
			'myDE.AddCustomLink "Interspire Website", "www.interspire.com", ""
		
			' Use the AddImageLibrary function to add an image library
			'myDE.AddImageLibrary "Image Library #1", "/devedit/flash_test"
			'myDE.AddImageLibrary "Image Library #2", "/images/library_1/other"
		
			' Use the AddFlashLibrary function to add flash libraries
			'myDE.AddFlashLibrary "Flash Library #1", "/flash_files"
			'myDE.AddFlashLibrary "Flash Library #2", "/flash_files/other"
		
			'Display the DevEdit control. This *MUST* be called between <form> and </form> tags
			'myDE.ShowControl "90%", "80%", "/DevEdit/ASP/images"
			myDE.ShowControl "90%", "400pt", "/timeout_xp/wwwroot/ver2_1/inc/upload/"&lto   '"/porsche/wwwroot/outsite/upload"
	end sub
	
	
	%>

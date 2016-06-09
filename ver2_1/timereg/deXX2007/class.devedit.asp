<%

	'/***********************************************************\
	'|                                                            |
	'|  DevEdit v5.0.0 Copyright Interspire Pty Ltd 2003		  |
	'|  All rights reserved. Do NOT modify this file. If		  |
	'|  you attempt to do so then we canot provide support. This  |
	'|  or any other DevEdit files may NOT be shared or			  |
	'|  distributed in any way. To purchase more licences, please |
	'|  visit www.devedit.com                                     |
	'|                                                            |
	'\************************************************************/

	'-------------------------------------
	'                                    '
	' DevEdit English Language Pack      '
	'                                    '
	'-------------------------------------

	Const sTxtOk = "OK"
	Const sTxtCancel = "Cancel"
	' Modified for v5.0
	Const sTxtCloseWin = "Click on the &quot;Cancel&quot; button to close this window."
	Const sTxtReturn = "Click 'Cancel' to return to the previous screen."
	Const sTxtName = "Name"
	Const sTxtBorder = "Border"
	Const sTxtBgColor = "Background Color"
	Const sTxtAction = "Action"
	Const sTxtRename = "Rename"
	Const sTxtBytes = "Bytes"
	Const sTxtFile = "File"
	Const sTxtTextAreaError = "Your browser must be IE5.5 or above to display the DevEdit control. A plain text box will be displayed instead."

	' Context Menu Width
	Const sTxtContextMenuWidth = 162

	' Added for v5.0
	' Color Window
	Const sTxtColors = "Colors"
	Const sTxtColorsInst = "Select your desired color and click 'OK' to use the selected color"

	' Find and Replace
	Const sTxtFindError = "Please enter text in the 'Find what:' field."
	Const sTxtFindNotFound = "Your word was not found. Would you like to start again from the top?"
	Const sTxtFindNotReplaced = "Word was not found. Nothing was replaced."
	Const sTxtFindReplaced = " word(s) were replaced."
	Const sTxtFindFindWhat = "Find what:"
	Const sTxtFindReplaceWith = "Replace with:"
	Const sTxtFindMatchCase = "Match case"
	Const sTxtFindMatchWord = "Match whole word only"
	Const sTxtFindNext = "Find Next"
	Const sTxtFindClose = "Close"
	Const sTxtFindReplaceText = "Replace"
	Const sTxtFindReplaceAll = "Replace All"

	' added for v5.0
	' flash
	Const sTxtInsertFlash = "Insert Flash"
	Const sTxtModifyFlash = "Modify Flash Properties"
	Const sTxtExternalFlash = "External Flash"
	Const sTxtInternalFlash = "Internal Flash"
	Const sTxtInsertFlashInst = "Enter the URL of the flash file to insert or select a flash file from those show below. Click 'Insert' to insert the flash file."
	Const sTxtUploadSuccess = " uploaded successfully"
	Const sTxtUploadFlash = "Upload Flash"
	Const sTxtFlashErr = "The file you uploaded was not a valid flash file"
	Const sTxtFlashExists = "already exists. Are you sure you want to overwrite it?"
	Const sTxtLoop = "Loop"
	Const sTxtChooseFlash = "Please chose a flash file to upload first!"

	Const sTxtCantDelete = "The selected file couldn't be deleted"
	Const sTxtCantUpload = "The selected file couldn't be uploaded"

	Const sTxtFlashWidthErr = "Flash width must contain a valid, positive number"
	Const sTxtFlashHeightErr = "Flash height must contain a valid, positive number"

	Const sTxtFlashWidth = "Flash Width"
	Const sTxtFlashHeight = "Flash Height"

	' End addition

	' Image
	Const sTxtInsertImage = "Insert Image"
	Const sTxtExternalImage = "External Image"
	Const sTxtInternalImage = "Internal Image"
	Const sTxtInsertImageInst = "Enter the URL of the image to insert or choose an image from those shown below and click on the insert button to add it to your content."
	Const sTxtUploadImage = "Upload Image"
	Const sTxtImageName = "File Name"
	Const sTxtFileSize = "File Size"
	Const sTxtImageView = "View"
	Const sTxtImageInsert = "Insert"
	Const sTxtInvalidImageType = "The selected file isn't a valid image"
	Const sTxtEmptyImageLibrary = "[The selected image library is empty]"

	' Added for v5.0
	Const sTxtImageModify = "Modify"
	' End addition

	Const sTxtImageDel = "Delete"
	Const sTxtImageBackgd = "Backgd"
	Const sTxtImageErr = "One of the files you uploaded was not a valid image file"
	Const sTxtImageSuccess = "Image(s) uploaded successfully"
	Const sTxtImageExists = "already exists. Are you sure you want to overwrite it?"
	Const sTxtChooseImage = "Please chose an image to upload first!"

	' Modified for v5.0
	Const sTxtImageDelete = "Are you sure you wish to delete this file?"
	Const sTxtImageDeleted = "Image deleted successfully!"
	' End modification

	Const sTxtImageSetBackgd = "Are you sure you wish to set this image as the page background image"

	' Link Manager
	Const sTxtLinkManager = "Link Manager"
	' Modified for v5.0
	Const sTxtLinkManagerInst = "Enter the required information and click on the &quot;OK&quot; button to insert a link into your webpage."
	Const sTxtLinkManagerInst2 = "Alternatively, locate the file from the file manager below and select &quot;Get Link Location&quot;. Click &quot;Insert Link&quot; to insert the link."
	Const sTxtLinkErr = "URL cannot be left blank"
	Const sTxtLinkInfo = "Link Information"
	Const sTxtUrl = "URL"

	'Added for v5.0
	Const sTxtLibraryUrl = "Pre-defined Links"
	Const sTxtTargetWin = "Target Window"
	Const sTxtAnch = "Anchor"
	Const sTxtGetLinkInfo = "Get link information to"
	Const sTxtGetLink = "Get Link Location"
	Const sTxtInsertLink = "Insert Link"
	Const sTxtRemoveLink = "Remove Link"
	' End addition

	' Anchor Global
	Const sTxtInsertAnchorErr = "Anchor name cannot be left blank"
	Const sTxtInsertAnchorName = "Anchor Name"

	' Anchor
	Const sTxtInsertAnchor = "Insert Anchor"
	' Modified for v5.0
	Const sTxtInsertAnchorInst = "Enter the required information and click on the &quot;OK&quot; button to insert an anchor into your webpage."

	' Modify Anchor
	Const sTxtModifyAnchor = "Modify Anchor"
	' Modified for v5.0
	Const sTxtModifyAnchorInst = "Enter the required information and click on the &quot;OK&quot; button to modify the current anchor."

	' Insert Button
	Const sTxtInsertButton = "Insert Button"
	Const sTxtInsertButtonInst = "Enter the required information and click on the &quot;OK&quot; button to insert a Button into your webpage."

	' Modify Button
	Const sTxtModifyButton = "Modify Button"
	Const sTxtModifyButtonInst = "Enter the required information and click on the &quot;OK&quot; button to modify the current button."

	' Insert Special Character
	Const sTxtInsertChar = "Insert Special Character"
	Const sTxtInsertCharInst = "Select the character you require and click on it to insert a it into your webpage."
	Const sTxtChar = "Character"
	Const sTxtCharToInsert = "Character to Insert"

	' Insert Checkbox
	Const sTxtInsertCheckbox = "Insert CheckBox"
	Const sTxtInsertCheckboxInst = "Enter the required information and click on the &quot;OK&quot; button to insert a checkBox into your webpage."

	' Modify Checkbox
	Const sTxtModifyCheckbox = "Modify CheckBox"
	Const sTxtModifyCheckboxInst = "Enter the required information and click on the &quot;OK&quot; button to modify the selected checkbox."

	' Insert Email Link
	Const sTxtEmailErr = "Email address cannot be left blank"
	Const sTxtInsertEmail = "Insert Email Link"
	Const sTxtRemoveEmailLink = "Remove Email Link"
	' Modified for v5.0
	Const sTxtInsertEmailInst = "Enter the required information and click on the &quot;OK&quot; button to create a link to an email address in your webpage."
	Const sTxtEmailAddress = "Email Address"
	Const sTxtSubject = "Subject"
	Const sTxtInsertEmailLink = "Insert Email Link"

	' Insert Hidden Field
	Const sTxtInsertHidden = "Insert Hidden Field"
	Const sTxtInsertHiddenInst = "Enter the required information and click on the &quot;OK&quot; button to insert a hidden field into your webpage."

	' Modify Hidden Field
	Const sTxtModifyField = "Modify Hidden Field"
	Const sTxtModifyFieldInst = "Enter the required information and click on the &quot;OK&quot; button to modify  the selected hidden field."

	' Insert Radio Button
	Const sTxtInsertRadio = "Insert Radio Button"
	Const sTxtInsertRadioInst = "Enter the required information and click on the &quot;OK&quot; button to insert a radio button into your webpage."

	' Modify Radio Button
	Const sTxtModifyRadio = "Modify Radio Button"
	Const sTxtModifyRadioInst = "Enter the required information and click on the &quot;OK&quot; button to modify the selected radio button."

	' Table Global
	Const sTxtTableRowErr = "Rows must contain a valid, positive number"
	Const sTxtTableColErr = "Columns must contain a valid, positive number"
	Const sTxtTableWidthErr = "Width must contain a valid, positive number"
	Const sTxtTablePaddingErr = "Cell padding must contain a valid, positive number"
	Const sTxtTableSpacingErr = "Cell spacing must contain a valid, positive number"
	Const sTxtTableBorderErr = "Border must contain a valid, positive number"

	Const sTxtRows = "Rows"
	Const sTxtPadding = "Cell Padding"
	Const sTxtSpacing = "Cell Spacing"
	Const sTxtWidth = "Width"
	Const sTxtHeight = "Height"
	Const sTxtCols = "Columns"

	' Insert Table
	Const sTxtInsertTable = "Insert Table"
	' Modified for v5.0
	Const sTxtInsertTableInst = "Enter the required information and click on the &quot;OK&quot; button to insert a table into your webpage."

	' Modify Table
	Const sTxtModifyTable = "Modify Table Properties"
	' Modified for v5.0
	Const sTxtModifyTableInst = "Enter the required information and click on the &quot;OK&quot; button to modify the properties of your table."

	' Modify Cell
	Const sTxtCellWidthErr = "Cell width must contain a valid, positive number"
	Const sTxtCellHeightErr = "Cell height must contain a valid, positive number"
	Const sTxtModifyCell = "Modify Cell Properties"
	' Modified for v5.0
	Const sTxtModifyCellInst = "Enter the required information and click on the &quot;OK&quot; button to modify the properties of your table cell."
	Const sTxtCellWidth = "Cell Width"
	Const sTxtCellHeight = "Cell Height"
	Const sTxtHorizontalAlign = "Horizontal Align"
	Const sTxtVerticalAlign = "Vertical Align"

	' Form Global
	Const sTxtCharWidthErr = "Character width must contain a valid, positive number"
	Const sTxtLinesErr = "Lines must contain a valid, positive number"
	Const sTxtMaxCharsErr = "Maximum characters must contain a valid, positive number"

	Const sTxtInitialValue = "Initial Value"
	Const sTxtInitialState = "Initial State"
	Const sTxtCharWidth = "Character Width"
	Const sTxtLines = "Lines"
	Const sTxtType = "Type"
	Const sTxtMaxChars = "Maximum Characters"

	Const sTxtMethod = "Method"

	' Insert Form
	Const sTxtInsertForm = "Insert Form"
	' Modified for v5.0
	Const sTxtInsertFormInst = "Enter the required information and click on the &quot;OK&quot; button to insert a form into your webpage."

	' Modify Form
	Const sTxtModifyForm = "Modify Form Properties"
	' Modified for v5.0
	Const sTxtModifyFormInst = "Enter the required information and click on the &quot;OK&quot; button to modify the properties of your form."

	' Insert TextArea
	Const sTxtInsertTextArea = "Insert Text Area"
	Const sTxtInsertTextAreaInst = "Enter the required information and click on the &quot;OK&quot; button to insert a text area into your webpage."

	' Modify TextArea
	Const sTxtModifyTextArea = "Modify Text Area"
	Const sTxtModifyTextAreaInst = "Enter the required information and click on the &quot;OK&quot; button to modify the selected text area."

	' Insert TextField
	Const sTxtInsertTextField = "Insert Text Field"
	' Modified for v5.0
	Const sTxtInsertTextFieldInst = "Enter the required information and click on the &quot;OK&quot; button to insert a text field into your webpage."

	' Modify TextField
	Const sTxtModifyTextField = "Modify Text Field"
	Const sTxtModifyTextFieldInst = "Enter the required information and click on the &quot;OK&quot; button to modify the selected text field."

	' JSCommon
	Const sTxtSetBackgd = "Are you sure you wish to set this image as the page background image?"

	' JSFunctions
	Const sTxtGuidelines = "Guidelines"
	Const sTxtOn = "ON"
	Const sTxtOff = "OFF"
	Const sTxtClean = "Are you sure you want to clean HTML code?"

	' Modify Image
	Const sTxtImageWidthErr = "Image width must contain a valid, positive number"
	Const sTxtImageHeightErr = "Image height must contain a valid, positive number"
	Const sTxtImageBorderErr = "Image border must contain a valid, positive number"
	Const sTxtHorizontalSpacingErr = "Horizontal spacing must contain a valid, positive number"
	Const sTxtVerticalSpacingErr = "Vertical spacing must contain a valid, positive number"

	Const sTxtModifyImage = "Modify Image Properties"
	Const sTxtModifyImageInst = "Enter the required information and click on the &quot;OK&quot; button to modify the properties of the selected image."
	Const sTxtAltText = "Alternate Text"
	Const sTxtImageWidth = "Image Width"
	Const sTxtImageHeight = "Image Height"
	Const sTxtAlignment = "Alignment"
	Const sTxtHorizontalSpacing = "Horizontal Spacing"
	Const sTxtVerticalSpacing = "Vertical Spacing"

	' Page Properties
	Const sTxtPageProps = "Modify Page Properties"
	Const sTxtPagePropsInst = "Enter the required information and click on the &quot;OK&quot; button to modify the properties of your page."

	Const sTxtPageTitle = "Page Title"
	Const sTxtDescription = "Description"
	Const sTxtKeywords = "Keywords"
	Const sTxtBgImage = "Background Image"
	Const sTxtTextColor = "Text Color"
	Const sTxtLinkColor = "Link Color"

	' Custom Inserts
	Const sTxtCustomInsert = "Custom Insert"

	' Modified for v5.0
	Const sTxtCustomInsertInst = "Select the required custom insert and click on the &quot;OK&quot; button"
	' End modification

	Const sTxtInsertHTML = "Insert HTML"
	Const sTxtCustomInsertErr = "Please select a custom insert to insert into your web page"
	Const sTxtPreview = "Preview"

	' Toolbar
	Const sTxtFullscreen = "Fullscreen Mode"
	Const sTxtCut = "Cut"
	Const sTxtCopy = "Copy"
	Const sTxtPaste = "Paste"
	Const sTxtFindReplace = "Find and Replace"
	Const sTxtUndo = "Undo"
	Const sTxtRedo = "Redo"
	Const sTxtRemoveFormatting = "Remove Text Formatting"
	Const sTxtBold = "Bold"
	Const sTxtUnderline = "Underline"
	Const sTxtItalic = "Italic"
	Const sTxtStrikethrough = "Strikethrough"
	Const sTxtNumList = "Insert Numbered List"
	Const sTxtBulletList = "Insert Bullet List"
	Const sTxtDecreaseIndent = "Decrease Indent"
	Const sTxtIncreaseIndent = "Increase Indent"
	Const sTxtSuperscript= "Superscript"
	Const sTxtSubscript= "Subscript"
	Const sTxtAlignLeft = "Align Left"
	Const sTxtAlignCenter = "Align Center"
	Const sTxtAlignRight = "Align Right"
	Const sTxtAlignJustify = "Justify"
	Const sTxtInsertHR = "Insert Horizontal Line"
	Const sTxtHyperLink = "Create or Modify Link"
	Const sTxtAnchor = "Insert / Modify Anchor"
	Const sTxtEmail = "Create Email Link"
	Const sTxtTextbox = "Insert Text Box"
	Const sTxtFormFunctions = "Form Functions"
	Const sTxtForm = "Insert Form"
	Const sTxtFormModify = "Modify Form Properties"
	Const sTxtTextField = "Insert / Modify Text field"
	Const sTxtTextArea = "Insert / Modify Text Area"
	Const sTxtHidden = "Insert / Modify Hidden Field"
	Const sTxtButton = "Insert / Modify Button"
	Const sTxtCheckbox = "Insert / Modify Checkbox"

	Const sTxtRadioButton = "Insert / Modify Radio Button"
	Const sTxtFont = "Font"
	Const sTxtSize = "Size"
	Const sTxtColor = "Color"
	Const sTxtFormat = "Format"
	Const sTxtStyle = "Style"
	Const sTxtColour = "Font Color"
	Const sTxtBackColour = "Highlight"
	Const sTxtTableFunctions = "Table Functions"
	Const sTxtTable = "Insert Table"
	Const sTxtTableModify = "Modify Table Properties"
	Const sTxtCellModify = "Modify Cell Properties"
	Const sTxtInsertRowA = "Insert Row Above"
	Const sTxtInsertRowB = "Insert Row Below"
	Const sTxtDeleteRow = "Delete Row"
	Const sTxtInsertColA = "Insert Column to the Right"
	Const sTxtInsertColB = "Insert Column to the Left"
	Const sTxtDeleteCol = "Delete Column"
	Const sTxtIncreaseColSpan = "Increase Column Span"
	Const sTxtDecreaseColSpan = "Decrease Column Span"
	Const sTxtImage = "Insert / Modify Image"

	'Added for v5.0
	Const sTxtFlash = "Insert / Modify Flash"
	Const sTxtFlashDeleted = "Flash file deleted successfully!"
	Const sTxtEmptyFlashLibrary = "[The selected flash library is empty]"

	Const sTxtChars = "Insert Special Characters"

	'Added for v5.0
	Const sTxtCharsInst = "Click on the required special character to insert that character into your webpage."

	Const sTxtPageProperties = "Modify Page Properties"
	Const sTxtCleanCode = "Clean HTML Code"
	Const sTxtPasteWord = "Paste from MS Word"
	Const sTxtCustomInserts = "Insert Custom HTML"
	Const sTxtToggleGuidelines = "Show / Hide Guidelines"
	Const sTxtTogglePosition = "Toggle Absolute Positioning"

	' Help
	Const sTxtHelpTitle = "&nbsp;The WYSIWYG Editor and commands"
	Const sTxtHelp = "Help"
	Const sTxtHelpNote = "Note: If an option below is not visible or accessible in your editor, then your administrator may have disabled it."
	Const sTxtHelpCloseWin = "Close Window"
	Const sTxtHelpSaveTitle = "Save"
	Const sTxtHelpSave = "To save your work, click on the 'Save' icon."
	Const sTxtHelpFullscreenTitle = "Fullscreen Mode"
	Const sTxtHelpFullscreen = "To expand the active window to a full screen click on the 'Fullscreen Mode' Icon. Concurrent clicks on this icon will toggle this feature on and off."
	Const sTxtHelpCutTitle = "Cut (Ctrl+X)"
	Const sTxtHelpCut = "To cut a portion of the document, highlight the desired portion and click the 'Cut' icon (keyboard shortcut - CTRL+X)."
	Const sTxtHelpCopyTitle = "Copy (Ctrl+C)"
	Const sTxtHelpCopy = "To copy a portion of the document, highlight the desired portion and click the 'Copy' icon (keyboard shortcut - CTRL+C)."
	Const sTxtHelpPasteTitle = "Paste (Ctrl+V)"
	Const sTxtHelpPaste = "To paste a portion that has already been cut (or copied), click where you want to place the desired portion on the page and click the 'Paste' icon (keyboard shortcut - CTRL+V).<br><br>To paste from Microsoft Word, click on the drop down icon next to the Paste Icon."
	Const sTxtHelpPasteWordTitle = "Paste from Microsoft Word (Ctrl + D)"
	Const sTxtHelpPasteWord = "To Paste from Microsoft Word: Copy your desired text from Microsoft Word and click the drop down icon next to the paste icon. Select the 'Paste from MS Word Option'. This will remove the tags that Microsoft Word automatically places around your text. It will also remove most text formatting as well."
	Const sTxtHelpFindReplaceTitle = "Find and Replace"
	Const sTxtHelpFindReplace = "To find and replace words or phrases within the text:<br> Select the search and replace feature. Enter the word or phrase you wish to replace and type it in the 'Find what' field<br><br>Select the new word or phrase you wish to replace the searched text with in the 'Replace with' field.<br><br>You can choose to 'find next' which allows you to manually replace instances of the searched text, or you can choose 'replace all' which allows you to replace all instances of the selected text.<br><br>Selecting the optional 'Match Case' tab allows you to search for a word or phrase with exactly the same upper or lower-case spelling of the word or phrase entered in 'Find What'. Not selecting this option means that a word entered in the 'Find what' field with upper case characters will return a search of upper and lower case matches of the same word.<br><br>Selecting the optional 'Match whole word only' tab allows the search to only display the words that are an exact match of the word or phrase entered in the 'Find What' field."
	Const sTxtHelpUndoTitle = "Undo (Ctrl+Z)"
	Const sTxtHelpUndo = "To undo the last change, click the 'Undo' icon (keyboard shortcut - CTRL+Z). Each consecutive click will undo the previous change to the document."
	Const sTxtHelpRedoTitle = "Redo (Ctrl+Y)"
	Const sTxtHelpRedo = "To redo the last change, click the 'Redo' icon (keyboard shortcut - CTRL+Y). Each consecutive click will repeat the last change to the document."
	Const sTxtHelpSpellcheckTitle = "Check Spelling (F7)"
	Const sTxtHelpSpellcheck = "To check spelling, select the text you would like to spell check (if you do not select the text, then your whole document will be checked)<br><br>Click on the spell checker icon or right click on the mouse and scroll down to 'Check spelling'.<br><br>You will be taken to the first incorrectly spelled word. You can then choose to<br><br>-	Change the incorrectly spelled word with the suggested words provided<br>-	Ignore the incorrectly spelled word (i.e. not make any changes to it)<br><br>To check spelling of a single word, highlight the word and right click on the mouse to get a selection of suggested replacements. To replace the miss-spelt word with one of the suggested words, simple select one of the replacements."
	Const sTxtHelpRemoveFormattingTitle = "Remove Text Formatting"
	Const sTxtHelpRemoveFormatting = "This command allows you to select a specific portion of text and remove any of the formatting which it contains. To remove any text formatting select the desired portion of text and Click the 'Remove Text Formatting' button."
	Const sTxtHelpBoldTitle = "Bold (Ctrl+B)"
	Const sTxtHelpBold = "To bold text, select the desired portion of text and click the 'Bold' icon (keyboard shortcut - CTRL+B). Each consecutive click will toggle this function on and off."
	Const sTxtHelpUnderlineTitle = "Underline (Ctrl+U)"
	Const sTxtHelpUnderline = "To underline text, select the desired portion of text and click the 'Underline' icon (keyboard shortcut - CTRL+U). Each consecutive click will toggle this function on and off."
	Const sTxtHelpItalicTitle = "Italic (Ctrl+I)"
	Const sTxtHelpItalic = "To convert text to italic, select the desired portion of text and click the 'Italic' icon (keyboard shortcut - CTRL+I). Each consecutive click will toggle this function on and off."
	Const sTxtHelpStrikethroughTitle = "Strikethrough"
	Const sTxtHelpStrikethrough = "To format text as strike through Select the text you want formatted by highlighting it and select the 'Strike through' icon. Each consecutive click will toggle this feature on and off"
	Const sTxtHelpINListTitle = "Insert Number List"
	Const sTxtHelpINList = "To start a numbered text list, click the 'Insert Numbered List' icon. If text has already been selected, the selection will be converted to a numbered list. Each consecutive click will toggle this function on and off."
	Const sTxtHelpIBListTitle = "Insert Bullet List"
	Const sTxtHelpIBList = "To start a bullet text list, click the 'Insert Bullet List' icon. If text has already been selected, the selection will be converted to a bullet list. Each consecutive click will toggle this function on and off."
	Const sTxtHelpDIndentTitle = "Decrease Indent"
	Const sTxtHelpDIndent = "To decrease indent of a paragraph, click the 'Decrease Indent' icon. Each consecutive click will move text further to the left."
	Const sTxtHelpIIndentTitle = "Increase Indent"
	Const sTxtHelpIIndent = "To increase indent of a paragraph, click the 'Increase Indent' icon. Each consecutive click will move text further to the right."
	Const sTxtHelpSuperscriptTitle = "Superscript"
	Const sTxtHelpSuperscript = "To convert text to superscript (vertically aligned higher): Select the desired portion of text and click the 'Superscript' icon.  Each consecutive click will toggle this function on and off."
	Const sTxtHelpSubscriptTitle = "Subscript"
	Const sTxtHelpSubscript = "To convert text to subscript (vertically aligned lower): Select the desired portion of text and click the 'Subscript' icon.  Each consecutive click will toggle this function on and off."
	Const sTxtHelpALeftTitle = "Align Left"
	Const sTxtHelpALeft = "To align to the left, make a selection in the document and click the 'Align Left' icon."
	Const sTxtHelpACenterTitle = "Align Center"
	Const sTxtHelpACenter = "To align to the center, make a selection in the document and click the 'Align Center' icon."
	Const sTxtHelpARightTitle = "Align Right"
	Const sTxtHelpARight = "To align to the right, make a selection in the document and click the 'Align Right' icon."
	Const sTxtHelpJustifyTitle = "Justify"
	Const sTxtHelpJustify = " To justify a paragraph or text, make a selection in the document and click the 'Justify' icon. "
	Const sTxtHelpIHLineTitle = "Insert Horizontal Line"
	Const sTxtHelpIHLine = "To insert a horizontal line, select the location to insert the line and click the 'Insert Horizontal Line' icon."
	Const sTxtHelpCMHyperLinkTitle = "Create or Modify Link"
	Const sTxtHelpCMHyperLink = "To create a hyperlink, select the text or image to create the link on, then click the 'Create or Modify Link' icon. if applicable, the popup window will contain existing link information. You can type the full URL of the page you want to link to in the URL text box. You can also enter the target window information (optional) and an anchor name (if linking to an anchor - optional).<br><br>For quick access to links, you can choose to insert a pre-defined link from the 'Pre-defined links' dropdown list.<br><br>When finished, click the 'Insert Link' button to insert the hyperLink you just created, or click 'Remove Link' to remove an existing link. Clicking 'Cancel' will close the window and take you back to the editor."
	Const sTxtHelpIMAnchorTitle = "Insert / Modify Anchor"
	Const sTxtHelpIMAnchor = "To insert an anchor, select a desired spot on the web page you are editing and click the 'Insert / Modify Anchor' icon. In the dialogue box, type the name for the anchor.<br><br>When finished, click the 'OK' button to insert the anchor, or 'Cancel' to close the box.<br><br>To modify an anchor select the anchor (displayed as an anchor icon when guidelines are switched on) and click the 'Insert / Modify Anchor' icon. Make your changes and hit the 'OK' button or click 'Cancel' to close the window"
	Const sTxtHelpCELinkTitle = "Create Email Link"
	Const sTxtHelpCELink = "To create an email link, select text or an image on the web page you are editing where you would like the link to appear. Click the 'Create Email Link' icon. In the dialogue box, type the email address for the link and the subject of the email.<br><br>When finished, click the 'OK' button to insert the email link, or 'Cancel' to close the box."
	Const sTxtHelpFontTitle = "Font"
	Const sTxtHelpFont = "To change the font of text, select the desired portion of text and click the 'Font' drop-down menu.<br><br>Select the desired font (choose from Default - Times New Roman, Arial, Verdana, Tahoma, Courier New or Georgia)."
	Const sTxtHelpFSizeTitle = "Font Size"
	Const sTxtHelpFSize = "To change the size of text, select the desired portion of text and click the 'Size' drop-down menu.<br><br>Select the desired size (text size 1-7)."
	Const sTxtHelpFormatTitle = "Format"
	Const sTxtHelpFormat = "To change the format of text, select the desired portion of text and click the 'Format' drop-down menu.<br><br>Select the desired format (choose from Normal, Superscript, Subscript and H1-H6)."
	Const sTxtHelpStyleTitle = "Style"
	Const sTxtHelpStyle = "To change the style of text, images, form objects, tables, table cells etc select the desired element and click the 'Style' drop-down menu, which will display all styles defined in the style-sheet.<br><br>Select the desired style from the menu.<br><br>Note: this function will not work if there is no style-sheet applied to the page."
	Const sTxtHelpFColorTitle = "Font Color"
	Const sTxtHelpFColor = "To change the colour of text, select the desired portion of text and click the 'Colour' drop-down menu.<br><br>Select the desired colour from the large selection in the drop-down menu. To define your own custom color, click on the 'More Colors...' button at the bottom of the color picker.<br><br>This will launch the advanced color picker, where you can select a color from the color map, or specify your own color using RGB or hex values. You can also change the contrast of the color by clicking on the contrast gradient"
	Const sTxtHelpHColorTitle = "Highlight Color"
	Const sTxtHelpHColor = "To change the highlighted colour of text, select the desired portion of text and click the 'Highlight' drop-down menu.<br><br>Select the desired colour from the large selection in the drop-down menu. To define your own custom color, click on the 'More Colors...' button at the bottom of the color picker.<br><br>This will launch the advanced color picker, where you can select a color from the color map, or specify your own color using RGB or hex values. You can also change the contrast of the color by clicking on the contrast gradient"
	Const sTxtHelpTFunctionsTitle = "Table Functions"
	Const sTxtHelpTFunctions = "To insert or modify a table or cell, select the 'Table Functions' icon to display a list of available Table Functions.<br><br>If a Table Function is NOT available, you will need to select, or place your cursor inside the table you wish to modify."
	Const sTxtHelpITableTitle = "Insert Table"
	Const sTxtHelpITable = "To insert a table, select the desired location, then click the 'Insert Table' icon.<br><br>A new window will pop-up with the following fields: Rows - number of rows in table; Columns - number of columns in table; Width - width of table; BgColour - background colour of table; Cell Padding - padding around cells; Cell Spacing - spacing between cells and Border - border around cells.<br><br>Fill in table details then click the 'OK' button to insert table, or click 'Cancel' to go back to the editor."
	Const sTxtHelpMTPropertiesTitle = "Modify Table Properties"
	Const sTxtHelpMTProperties = "To modify table properties, select a table or click anywhere inside the table to modify, then click the 'Modify Table Properties' icon.<br><br>A pop-up window will appear with the table's properties. Click the 'OK' button to save your changes, or click 'Cancel' to go back to the editor.<br><br>Note: this function will not work if a table has not been selected."
	Const sTxtHelpMCPropertiesTitle = "Modify Cell Properties"
	Const sTxtHelpMCProperties = "To modify cell properties, click inside the cell to modify, then click the 'Modify Cell Properties' icon.<br><br>A pop-up window will appear with the cells' properties.<br><br>Click the 'OK' button to save your changes, or click 'Cancel' to go back to the editor.<br><br>Note: this function will not work if a cell has not been selected and does not work across multiple cells."
	Const sTxtHelpICttRightTitle = "Insert Column to the Right"
	Const sTxtHelpICttRight = "To insert a column to the right of your cursor, click inside cell after which to insert a column, then click the 'Insert Column to the Right' icon.<br><br>Each consecutive click will insert another column after the selected cell.<br><br>Note: this function will not work if a cell has not been selected."
	Const sTxtHelpICttLeftTitle = "Insert Column to the Left"
	Const sTxtHelpICttLeft = "To insert column to the left of your cursor, click inside cell before which to insert a column, then click the 'Insert Column to the Left' icon.<br><br>Each consecutive click will insert another column before the selected cell.<br><br>Note: this function will not work if a cell has not been selected."
	Const sTxtHelpIRAboveTitle = "Insert Row Above"
	Const sTxtHelpIRAbove = "To insert row above, click inside cell above which to insert a row, then click the 'Insert Row Above' icon.<br><br>Each consecutive click will insert another row above the selected cell.<br><br>Note: this function will not work if a cell has not been selected."
	Const sTxtHelpIRBelowTitle = "Insert Row Below"
	Const sTxtHelpIRBelow = "To insert row below, click inside cell below which to insert a row, then click the 'Insert Row Below' icon.<br><br>Each consecutive click will insert another row below the selected cell.<br><br>Note: this function will not work if a cell has not been selected."
	Const sTxtHelpDRowTitle = " Delete Row"
	Const sTxtHelpDRow = "To delete a row, click inside cell which is in the row to be deleted, then click the 'Delete Row' icon.<br><br>Note: this function will not work if a cell has not been selected."
	Const sTxtHelpIColumnTitle = "Insert Column"
	Const sTxtHelpIColumn = "To insert a column, click inside cell which is in the column to be inserted, then click the 'Insert Column' icon.<br><br>Note: this function will not work if a cell has not been selected."
	Const sTxtHelpDColumnTitle = "Delete Column"
	Const sTxtHelpDColumn = "To delete a column, click inside cell which is in the column to be deleted, then click the 'Delete Column' icon.<br><br>Note: this function will not work if a cell has not been selected."
	Const sTxtHelpICSpanTitle = "Increase Column Span"
	Const sTxtHelpICSpan = "To insert column span, click inside cell who's span is to be increased, then click the 'Increase Column Span' icon.<br><br>Each consecutive click will further increase the column span of the selected cell.<br><br>Note: this function will not work if a cell has not been selected."
	Const sTxtHelpDCSpanTitle = "Decrease Column Span"
	Const sTxtHelpDCSpan = "To decrease column span, click inside cell who's span is to be decreased, then click the 'Decrease Column Span' icon.<br><br>Each consecutive click will further decrease the column span of the selected cell. Note: this function will not work if a cell has not been selected."
	Const sTxtHelpFFunctionsTitle = "Form Functions"
	Const sTxtHelpFFunctions = "To insert or modify a form, select the 'Form Functions' icon to display a list of available form functions.<br><br>If a form function is NOT available, you will need to place your cursor inside the form you wish to modify."
	Const sTxtHelpIFormTitle = "Insert Form"
	Const sTxtHelpIForm = "To insert a form, select desired position then click the 'Insert Form' icon.<br><br>A new window will pop-up with the following fields: Name - name of form; Action - location of script that processes the form and Method - post, get or none.<br><br>Fill in form details or leave blank for a blank form.<br><br>When finished, click the 'OK' button to insert form, or click 'Cancel' to go back to the editor."
	Const sTxtHelpMFPropertiesTitle = "Modify Form Properties"
	Const sTxtHelpMFProperties = "To modify form properties, click anywhere inside the form to modify, then click the 'Modify Form Properties' icon.<br><br>A pop-up window will appear with the form's properties.<br><Br>Click the 'OK' button to save your changes, or click 'Cancel' to go back to the editor. Note: this function will not work if a form has not been selected."
	Const sTxtHelpIMTFieldTitle = "Insert / Modify Text Field"
	Const sTxtHelpIMTField = "To insert a text field, select the desired position then click the 'Insert/Modify Text Field' icon.<br><br>A pop-up window will appear with the following attributes: Name - name of text field; Character width - the width of the text field, in characters; Type - type of text field (Text or Password); Initial value - initial text in field and Maximum characters - maximum number of characters allowed.<br><br>Set the attributes then click the 'OK' button to insert text field, or click 'Cancel' to go back to the editor.<br><br>To modify a text field's properties, select desired text field and click the 'Insert/Modify Text Field' icon.<br><Br>A pop-up window will appear with the text field's attributes.<br><br>Modify any attributes desired, then click the 'OK' button to save changes, or click 'Cancel' to go back to the editor."
	Const sTxtHelpIMTAreaTitle = "Insert / Modify Text Area"
	Const sTxtHelpIMTArea = "To insert a text area, select the desired position then click the 'Insert/Modify Text Area' icon<br><br>A pop-up window will appear with the following attributes: Name - name of text area; Character width - the width of the text area, in characters; Initial value - initial text in area and Lines - number of lines allowed in the text area.<br><br>Set the attributes then click the 'OK' button to insert the text area, or click 'Cancel' to go back to the editor.<br><br>To modify a text area's properties, select desired text area and click the 'Insert/Modify Text Area' icon.<br><br>A pop-up window will appear with the text area's attributes.<br><br>Modify any attributes desired, then click the 'OK' button to save changes, or click 'Cancel' to go back to the editor. "
	Const sTxtHelpIMHAreaTitle = "Insert / Modify Hidden Area"
	Const sTxtHelpIMHArea = "To insert a hidden field, select desired position then click the 'Insert/Modify Hidden Field' icon.<br><br>A pop-up window will appear with the following attributes: Name - name of hidden field and Initial value - initial value of hidden field.<br><br>Set the attributes then click the 'OK' button to insert the hidden field, or click 'Cancel' to go back to the editor.<br><br>To modify a hidden field's properties, select desired hidden field and click the 'OK' icon.<br><br>A pop-up window will appear with the hidden field's attributes.<br><br>Modify any attributes desired, then click the 'OK' button to save changes or click 'Cancel' to go back to the editor. "
	Const sTxtHelpIMButtonTitle = "Insert / Modify Button"
	Const sTxtHelpIMButton = "To insert a button, select desired position then click the 'Insert/Modify Button' icon.<br><br>A pop-up window will appear with the following attributes: Name - name of text area; Type - type of button (Submit, Reset or Button) and Initial value - initial text on the button.<br><br>Set the attributes then click 'OK' to insert button, or click 'Cancel' to go back to the editor.<br><br>To modify a button's properties, select desired button and click the 'Insert/Modify Button' icon.<br><br>A pop-up window will appear with the button's attributes.<br><br>Modify any attributes desired, then click the 'OK' button to save changes or click 'Cancel' to go back to the editor. "
	Const sTxtHelpIMCheckboxTitle = "Insert / Modify Checkbox"
	Const sTxtHelpIMCheckbox = "To insert a checkbox, select desired position then click the 'Insert/Modify Checkbox' icon.<br><br>A pop-up window will appear with the following attributes: Name - name of checkbox; Initial state - checked or unchecked and Initial value - value of checkbox.<br><br>Set the attributes then click the 'OK' button to insert the checkbox, or click 'Cancel' to go back to the editor.<br><br>To modify a checkbox' properties, select desired checkbox and click the 'Insert/Modify Checkbox' icon.<br><br>A pop-up window will appear with the checkbox' attributes.<br><br>Modify any attributes desired, then click the 'OK' button to save changes or click 'Cancel' to go back to the editor. "
	Const sTxtHelpIMRButtonTitle = "Insert / Modify Radio Button"
	Const sTxtHelpIMRButton = "To insert a radio button, select desired position then click the 'Insert/Modify Radio Button' icon.<br><br>A pop-up window will appear with the following attributes: Name - name of radio button; Initial state - checked or unchecked and Initial value - value of radio button.<br><br>Set the attributes then click 'OK' to insert the radio button, or click 'Cancel' to go back to the editor.<br><br>To modify a checkbox' properties, select desired checkbox and click the 'Insert/Modify Radio Button' icon.<br><br>A pop-up window will appear with the checkbox' attributes.<br><br>Modify any attributes desired, then click the 'OK' button to save changes or click 'Cancel' to go back to the editor. "
	
	Const sTxtHelpIMFlashTitle = "Insert / Modify Flash"
	Const sTxtHelpIMFlash = "If a flash movie is NOT selected, clicking this icon will open the Flash Manager.<br><br>If a flash movie IS selected, then clicking this icon will open the 'Modify Flash Properties' popup window.<br><br>To modify the properties of the selected flash movie, set the required attributes and click the 'Modify' button."

	Const sTxtHelpIMImageTitle = "Insert / Modify Image"
	Const sTxtHelpIMImage = "If an image is NOT selected, clicking this icon will open the Image Manager.<br><br>If an image IS selected, then clicking this icon will open the 'Modify Image Properties' popup window.<br><br>To modify the image properties of the selected image, set the required attributes and click the 'Modify' button."

	Const sTxtHelpTextboxTitle = "Insert Textbox"
	Const sTxtHelpTextbox = "To add a text box anywhere within the page, select the location where you want the text box to reside in the active window and click on the 'insert text box icon' that will place a text box where you have specified. <br><br>To resize the text box, click on the text box frame (turn 'show/hide guidelines' on if you cannot see the textbox outline). Then click on side/corner of the frame you wish to resize from and drag to a size you require. The text you type will be contained within the text box and will stretch to the size of the text box."

	Const sTxtHelpAbsoluteTitle = "Toggle Absolute Positioning"
	Const sTxtHelpAbsolute = "To position a text box or image using absolute positioning, select the the textbox or image and select the 'absolute positioning' icon. You can now click and drag an image or text box to the location you would like it to reside within the active window."

	Const sTxtHelpISCharactersTitle = "Insert Special Characters"
	Const sTxtHelpISCharacters = "To insert a special character, click the 'Insert Special Character' icon.<br><br>A pop-up window will appear with a list of special characters.<br><br>Click the icon of the character to insert into your webpage."
	Const sTxtHelpMPPropertiesTitle = "Modify Page Properties"
	Const sTxtHelpMPProperties = "To modify page properties, click the 'Modify Page Properties' icon.<br><br>A pop-up window will appear with page attributes: Page Title - title of page; Description - description of page; Background Image - The URL of the image curently set as the page background image; Keywords - keywords page is to be recognized by; Background Colour - the background colour of page; Text Colour - colour of text in page and Link Colour - the colour of links in page.<br><br>When finished modifying, click the 'OK' button to save changes, or click 'Cancel' to go back to the editor."
	Const sTxtHelpCUHTMLCodeTitle = " Clean Up HTML Code"
	Const sTxtHelpCUHTMLCode = "To clean HTML code, click the 'Clean HTML Code' icon.<br><br>This will remove any empty span and paragraph tags, all xml tags, all tags that have a colon in the tag name (i.e. <tag:with:colon>) and remove style and class attributes.<br><br>This is useful when copying and pasting from Microsoft Word documents to remove unnecessary HTML code. "
	Const sTxtHelpCustomHTMLTitle = "Insert Custom HTML"
	Const sTxtHelpCustomHTML = "There may be a list of available items to insert that you can preview and choose from. This list will usually contain customized items in HTML such as logos and formatted text specific to your site. To preview an item, click on the item in the list, and the item will appear the preview field below. To select the item, click on it and choose 'OK'."
	Const sTxtHelpSHGuidelinesTitle = "Show / Hide Guidelines"
	Const sTxtHelpSHGuidelines = "To show or hide guidelines, click the 'Show/Hide Guidelines' icon.<br><br>This will toggle between displaying table and form guidelines and not showing any guidelines at all.<br><br>Tables and cells will have a broken grey line around them, forms will have a broken red line around them, while hidden fields will be a pink square when showing guidelines.<br><br>Note that the status bar (at the bottom of the window) will reflect the guidelines mode currently in use."
	Const sTxtHelpSModeTitle = "Source Mode"
	Const sTxtHelpSMode = " To switch to source code editing mode, click the 'Source' button at the bottom of the editor.<br><br>This will switch to HTML code editng mode.<br><br>To switch back to WYSIWYG Editing, click the 'Edit' button at the bottom of the editor.<br><br>This is recommended for advanced users only "
	Const sTxtHelpPModeTitle = "Preview Mode"
	Const sTxtHelpPMode = "To show a preview of the page being edited, click the 'Preview' button at the bottom of the editor.<br><br>This is useful in previewing a file to see how it will look exactly in your browser, before changes are saved.<br><br>To switch back to editing mode, click the 'Edit' button at the bottom of the editor."
	Const sTxtHelpImageManager = "Image Manager"
	Const sTxtHelpBTTOP = "Back To Top"
	Const sTxtHelpImageDescription = "The Image Manager is where you can preview, insert, delete and upload your image files.<br><br>You can perform general maintenance on your images from the Image Manager - insert, set as background, upload, view, and delete images."
	Const sTxtHelpVaImageTitle = "Viewing an Image"
	Const sTxtHelpVaImage = "To view an image, select the desired image and click on the 'preview' link.<br><br>The image will be shown in a pop-up window in it's full size.<br>Close the window to return to the Image Manager."
	Const sTxtHelpIaImageTitle = "Inserting an Image"
	Const sTxtHelpIaImage = "To insert an image, click the 'Insert' link at the bottom of the image manager."
	Const sTxtHelpSBImageTitle = "Set Background Image"
	Const sTxtHelpSBImage = "To set an image as a background image, click the 'Backgd' button in the image browser. <br><br>The image will be set as the current page background image."
	Const sTxtHelpDaImageTitle = "Deleting an Image"
	Const sTxtHelpDaImage = "To delete, select the desired image and click on the 'Delete' button.<br><br>You will be prompted for confirmation of the deletion.<br><br>If you are sure you wish to delete the selected image, click 'OK'.<br><br>Clicking on 'Cancel' will take you back to the Image Manager."
	Const sTxtHelpUaImageTitle = "Uploading an Image"
	Const sTxtHelpUaImage = "To upload an image, click the 'Browse' button to open a 'Choose File' box that allows you to select a local image to upload.<br><br>Once the file has been selected, click 'OK' to begin uploading the file, or click 'Cancel' to be taken back to the Image Manager<br><br>Upon sucessful upload of the image, it will appear in the Image Manager.<br><br>To upload more than one image, click on the » button. You can upload up to 5 images at any one time."

	' Added for v5.0.1 fix
	Const sTxtSave = "Save"
	Const sTxtNone = "None"
	Const sTxtMoreColors = "More Colors"
	Const sTxtCheckSpelling = "Check Spelling"
	Const sTxtUpload = "Upload"
	Const sTxtSwitch = "Switch"
	Const sTxtImageDirNotConfigured = "Image directory has not been configured correctly"
	Const sTxtFlashDirNotConfigured = "Flash directory has not been configured correctly"
	Const sTxtImageProperties = "Image Properties"
	Const sTxtFlashProperties = "Flash Properties"
	Const sTxtDefaultImageLibrary = "Default Image Library"
	Const sTxtDefaultFlashLibrary = "Default Flash Library"
	Const sTxtLoad = "Load"

	'Added for v5.0.1 fix
	' Insert Select
	Const sTxtInsertSelect = "Insert Select"
	Const sTxtInsertSelectInst = "Enter the required information and click on the &quot;OK&quot; button to insert a select field into your webpage."
	Const sTxtText = "Text"
	Const sTxtValue = "Value"
	Const sTxtSelected = "Selected"
	Const sTxtSingle = "Single Select"
	Const sTxtMultiple = "Multiple Select"
	Const sTxtAdd = "Add"
	Const sTxtUpdate = "Update"
	Const sTxtDelete = "Delete"
	Const sTxtMaintainOptions = "Maintain Options"
	Const sTxtCurrentOptions = "Current Options"

	Const sTxtModifySelect = "Modify Select"
	Const sTxtModifySelectInst = "Enter the required information and click on the &quot;OK&quot; button to modify the selected select field."

	Const sTxtSelect = "Insert / Modify Select Field"
	Const sTxtHelpIMSListTitle = "Insert / Modify Select Field"
	Const sTxtHelpIMSList = "To insert a select field, select the desired position then click the 'Insert/Modify Select Field' icon.<br><br>A pop-up window will appear with the following attributes: Name - name of the select list; Current Options - The options available for selection in the list; Type - how the list will be displayed (a single option, or multiple options); Size - how many list items will be shown; Style - The style to be applied to this select field, if any.<br><br>To add options to the select list, use the text, value and selected boxes under the 'Maintain Options' heading.<br><br>To modify a select lists properties, select the desired list and click the 'Insert/Modify Select List' button. A popup window will appear with the select lists attributes.<br><br>Modify the desired attributes, then click the 'OK' button to save changes or click 'Cancel' to go back to the editor."
	
	' End addition

	'Constant variables to make function calling more logical
	const de_PATH_TYPE_FULL = 0
	const de_PATH_TYPE_ABSOLUTE = 1
	const de_DOC_TYPE_SNIPPET = 0
	const de_DOC_TYPE_HTML_PAGE = 1
	const de_IMAGE_TYPE_ROW = 0
	const de_IMAGE_TYPE_THUMBNAIL = 1
	const de_FLASH_TYPE_ROW = 0
	const de_FLASH_TYPE_THUMBNAIL = 1

	'Language variations
	const de_AMERICAN = 1
	const de_BRITISH = 2
	const de_CANADIAN = 3
	const de_FRENCH = 4
	const de_SPANISH = 5
	const de_GERMAN = 6
	const de_ITALIAN = 7
	const de_PORTUGESE = 8
	const de_DUTCH = 9
	const de_NORWEGIAN = 10
	const de_SWEDISH = 11
	const de_DANISH = 12

	' Check if HTTPS is enabled
	Dim HTTPStr
	Dim DevEditPath
	Dim DevEditError

	if UCase(Request.ServerVariables("HTTPS")) = "ON" then
		HTTPStr = "https"
	else
		HTTPStr = "http"
	end if

	public sub SetDevEditPath(Path)

		dim tmpPath, lastChar, firstSlash

		tmpPath = ""

		'Does the path contain a trailing slash? If so, remove it
		lastChar = right(Path, 1)

		if lastChar = "/" then
			tmpPath = left(Path, len(Path)-1)
		else
			tmpPath = Path
		end if

		'Is this a relative path?
		if left(tmpPath, 1) <> "/" then
			tmpPath = strreverse(Request.ServerVariables("SCRIPT_NAME"))
			firstSlash = instr(tmpPath, "/")
			tmpPath = strreverse(mid(tmpPath, firstSlash, len(tmpPath)))
			tmpPath = tmpPath & "/" & Path
			tmpPath = replace(tmpPath, "//", "/")
		end if

		DevEditPath = tmpPath

	end sub

	public sub DisplayIncludes (file, errorMsg)

		Const ForReading = 1, ForWriting = 2, ForAppending = 8 

		dim fso, f, ts, fileContent, includeFile
		dim URL, scriptName, serverName, scriptDir, slashPos, oMatch
		set fso = server.CreateObject("Scripting.FileSystemObject")

		if DevEditPath = "" then
			DevEditPath = Request("DEP")
		end if

		includeFile = Server.mapPath(DevEditPath & "/de_includes/" & file)

		if(Request.QueryString("DEP1") <> "") then
			deveditPath = Request.QueryString("DEP1")
		end if

		if (fso.FileExists(includeFile)=true) Then

			set f = fso.GetFile(includeFile)
			set ts = f.OpenAsTextStream(ForReading, -2) 
				
			Do While not ts.AtEndOfStream
				fileContent = fileContent & ts.ReadLine & vbCrLf
			Loop

			URL = Request.ServerVariables("http_host")
			scriptName = DevEditPath & "/class.devedit.asp"

			'Workout the location of class.devedit.asp
			scriptDir = strreverse(Request.ServerVariables("path_info"))
			slashPos = instr(1, scriptDir, "/")
			scriptDir = strreverse(mid(scriptDir, slashPos, len(scriptDir)))

			scriptName = scriptDir & scriptName
				
			fileContent = replace(fileContent,"$URL",URL)
			fileContent = replace(fileContent,"$SCRIPTNAME",ScriptName)
			fileContent = replace(fileContent,"$HTTPStr",HTTPStr)

			'added for v5.0
			fileContent = replace(fileContent,"$DEP",deveditPath)
				
			Dim re
			Set re = New RegExp
			re.global = true

			re.Pattern = "\[sTxt(\w*)\]"

			For Each oMatch in re.Execute(fileContent)
			 	fileContent = replace(fileContent,oMatch,eval("sTxt" & oMatch.SubMatches(0)))
			Next

			response.write(fileContent)

		else
			response.write("file not found:" & file)
			DevEditError = true
		End if
		
	End Sub

	' Examine the value of the ToDo argument and proceed to correct sub
	dim ToDo
	ToDo = Request("ToDo")

	if ToDo = "" then
	%>
		<link rel="stylesheet" href="<%=DevEditPath%>/de_includes/de_styles.css" type="text/css">
	<%
	end if

	if ToDo = "InsertImage" Then
		' pass to insert image screen
		PageInsertImage()
	elseif ToDo = "DeleteImage" Then
		PageInsertImage()
	elseif ToDo = "UploadImage" Then
		PageInsertImage()
	elseif ToDo = "InsertFlash" Then
		PageInsertFlash()
	elseif ToDo = "UploadFlash" Then
		PageInsertFlash()
	elseif ToDo = "DeleteFlash" Then
		PageInsertFlash()
	elseif ToDo = "FindReplace" Then
		DisplayIncludes "find_replace.inc","Find and Replace"
	elseif ToDo = "SpellCheck" Then
		DisplayIncludes "spell_check.inc","Spell Check"
	elseif ToDo = "DoSpell" Then
		DisplayIncludes "do_spell.inc","Spell Check"
	elseif ToDo = "InsertTable" Then
		DisplayIncludes "insert_table.inc","Insert Table"
	elseif ToDo = "ModifyTable" Then
		DisplayIncludes "modify_table.inc","Modify Table"
	elseif ToDo = "ModifyCell" Then
		DisplayIncludes "modify_cell.inc","Modify Cell"
	elseif ToDo = "ModifyImage" Then
		DisplayIncludes "modify_image.inc","Modify Image"
	elseif ToDo = "InsertForm" Then
		DisplayIncludes "insert_form.inc","Insert Form"
	elseif ToDo = "ModifyForm" Then
		DisplayIncludes "modify_form.inc","Modify Form"
	elseif ToDo = "InsertTextField" Then
		DisplayIncludes "insert_textfield.inc","Insert Text Field"
	elseif ToDo = "ModifyTextField" Then
		DisplayIncludes "modify_textfield.inc","Modify Text Field"
	elseif ToDo = "InsertTextArea" Then
		DisplayIncludes "insert_textarea.inc","Insert Text Area"
	elseif ToDo = "ModifyTextArea" Then
		DisplayIncludes "modify_textarea.inc","Modify Text Area"
	elseif ToDo = "InsertHidden" Then
		DisplayIncludes "insert_hidden.inc","Insert Hidden Field"
	elseif ToDo = "ModifyHidden" Then
		DisplayIncludes "modify_hidden.inc","Modify Hidden Field"
	elseif ToDo = "InsertButton" Then
		DisplayIncludes "insert_button.inc","Insert Button"
	elseif ToDo = "ModifyButton" Then
		DisplayIncludes "modify_button.inc","Modify Button"
	elseif ToDo = "InsertCheckbox" Then
		DisplayIncludes "insert_checkbox.inc","Insert Checkbox"
	elseif ToDo = "ModifyCheckbox" Then
		DisplayIncludes "modify_checkbox.inc","Modify CheckBox"
	elseif ToDo = "InsertSelect" Then
		DisplayIncludes "insert_select.inc","Insert Select"
	elseif ToDo = "ModifySelect" Then
		DisplayIncludes "modify_select.inc","Modify Select"
	elseif ToDo = "InsertRadio" Then
		DisplayIncludes "insert_radio.inc","Insert Radio"
	elseif ToDo = "ModifyRadio" Then
		DisplayIncludes "modify_radio.inc","Modify Radio"
	elseif ToDo = "PageProperties" Then
		DisplayIncludes "page_properties.inc","Page Properties"
	elseif ToDo = "Chars" Then
		DisplayIncludes "insert_chars.inc","Insert Special Characters"
	elseif ToDo = "InsertLink" Then
		DisplayIncludes "insert_link.inc","Insert Link"
	elseif ToDo = "InsertEmail" Then
		DisplayIncludes "insert_email.inc","Insert Email Link"
	elseif ToDo = "InsertAnchor" Then
		DisplayIncludes "insert_anchor.inc","Insert Anchor"
	elseif ToDo = "ModifyAnchor" Then
		DisplayIncludes "modify_anchor.inc","Modify Anchor"
	elseif ToDo = "MoreColors" Then
		DisplayIncludes "more_colors.inc","More Colors"
	elseif ToDo = "CustomInsert" Then
		DisplayIncludes "custom_insert.inc","Insert Custom HTML"
	elseif ToDo = "ShowHelp" Then
		DisplayIncludes "help.inc","Help"
	End if

	  ' -- Loader.asp --
	  ' -- version 1.5
	  ' -- last updated 6/13/2002
	  '
	  ' Faisal Khan
	  ' faisal@stardeveloper.com
	  ' www.stardeveloper.com
	  ' Class for handling binary uploads
	  
	  On Error Resume Next
	  
	  dim i, tPoint
	  
	  Class Loader
		Private dict
		
		Private Sub Class_Initialize
		  Set dict = Server.CreateObject("Scripting.Dictionary")
		End Sub
		
		Private Sub Class_Terminate
		  If IsObject(intDict) Then
			intDict.RemoveAll
			Set intDict = Nothing
		  End If
		  If IsObject(dict) Then
			dict.RemoveAll
			Set dict = Nothing
		  End If
		End Sub

		Public Property Get Count
		  Count = dict.Count
		End Property

		Public Sub Initialize
		  If Request.TotalBytes > 0 Then
			Dim binData
			  binData = Request.BinaryRead(Request.TotalBytes)
			  getData binData
		  End If
		End Sub
		
		Public Function getFileData(name)
		  If dict.Exists(name) Then
			getFileData = dict(name).Item("Value")
			Else
			getFileData = ""
		  End If
		End Function

		Public Function getValue(name)
		  Dim gv
		  If dict.Exists(name) Then
			gv = CStr(dict(name).Item("Value"))
			
			gv = Left(gv,Len(gv)-2)
			getValue = gv
		  Else
			getValue = ""
		  End If
		End Function
		
		Public Function saveToFile(name, path)
		  If dict.Exists(name) Then
			Dim temp
			  temp = dict(name).Item("Value")
			Dim fso
			  Set fso = Server.CreateObject("Scripting.FileSystemObject")
			Dim file
			  Set file = fso.CreateTextFile(path)
				For tPoint = 1 to LenB(temp)
					file.Write Chr(AscB(MidB(temp,tPoint,1)))
				Next
				file.Close
			  saveToFile = True
		  Else
			  saveToFile = False
		  End If
		End Function
		
		Public Function getFileName(name)
		  If dict.Exists(name) Then
			Dim temp, tempPos
			  temp = dict(name).Item("FileName")
			  tempPos = 1 + InStrRev(temp, "\")
			  getFileName = Mid(temp, tempPos)
		  Else
			getFileName = ""
		  End If
		End Function
		
		Public Function getFilePath(name)
		  If dict.Exists(name) Then
			Dim temp, tempPos
			  temp = dict(name).Item("FileName")
			  tempPos = InStrRev(temp, "\")
			  getFilePath = Mid(temp, 1, tempPos)
		  Else
			getFilePath = ""
		  End If
		End Function
		
		Public Function getFilePathComplete(name)
		  If dict.Exists(name) Then
			getFilePathComplete = dict(name).Item("FileName")
		  Else
			getFilePathComplete = ""
		  End If
		End Function

		Public Function getFileSize(name)
		  If dict.Exists(name) Then
			getFileSize = LenB(dict(name).Item("Value"))
		  Else
			getFileSize = 0
		  End If
		End Function

		Public Function getFileSizeTranslated(name)
		  If dict.Exists(name) Then
			temp = 1 + LenB(dict(name).Item("Value"))
			  If Len(temp) <= 3 Then
				getFileSizeTranslated = temp & " bytes"
			  ElseIf Len(temp) > 6 Then
				temp = FormatNumber(((temp / 1024) / 1024), 2)
				getFileSizeTranslated = temp & " megabytes"  
			  Else
				temp = FormatNumber((temp / 1024), 2)
				getFileSizeTranslated = temp & " kilobytes"
			  End If
		  Else
			getFileSizeTranslated = ""
		  End If
		End Function
		
		Public Function getContentType(name)
		  If dict.Exists(name) Then
			getContentType = dict(name).Item("ContentType")
		  Else
			getContentType = ""
		  End If
		End Function

	  Private Sub getData(rawData)
		Dim separator 
		  separator = MidB(rawData, 1, InstrB(1, rawData, ChrB(13)) - 1)

		Dim lenSeparator
		  lenSeparator = LenB(separator)

		Dim currentPos
		  currentPos = 1
		Dim inStrByte
		  inStrByte = 1
		Dim value, mValue
		Dim tempValue
		  tempValue = ""

		While inStrByte > 0
		  inStrByte = InStrB(currentPos, rawData, separator)
		  mValue = inStrByte - currentPos

		  If mValue > 1 Then
			value = MidB(rawData, currentPos, mValue)

			Dim begPos, endPos, midValue, nValue
			Dim intDict
			  Set intDict = Server.CreateObject("Scripting.Dictionary")
		
			  begPos = 1 + InStrB(1, value, ChrB(34))
			  endPos = InStrB(begPos + 1, value, ChrB(34))
			  nValue = endPos

			Dim nameN
			  nameN = MidB(value, begPos, endPos - begPos)

			Dim nameValue, isValid
			  isValid = True
			  
			  If InStrB(1, value, stringToByte("Content-Type")) > 1 Then

				begPos = 1 + InStrB(endPos + 1, value, ChrB(34))
				endPos = InStrB(begPos + 1, value, ChrB(34))
	  
				If endPos = 0 Then
				  endPos = begPos + 1
				  isValid = False
				End If
				
				midValue = MidB(value, begPos, endPos - begPos)
				  intDict.Add "FileName", trim(byteToString(midValue))
					
				begPos = 14 + InStrB(endPos + 1, value, stringToByte("Content-Type:"))
				endPos = InStrB(begPos, value, ChrB(13))
				
				midValue = MidB(value, begPos, endPos - begPos)
				  intDict.Add "ContentType", trim(byteToString(midValue))
				
				begPos = endPos + 4
				endPos = LenB(value)
				
				nameValue = MidB(value, begPos, endPos - begPos)
			  Else
				nameValue = trim(byteToString(MidB(value, nValue + 5)))
			  End If

			  If isValid = true Then
				intDict.Add "Value", nameValue
				intDict.Add "Name", nameN

				dict.Add byteToString(nameN), intDict
			  End If
		  End If

		  currentPos = lenSeparator + inStrByte
		Wend
	  End Sub
	  
	  End Class

	  Private Function stringToByte(toConv)
		Dim tempChar
		 For i = 1 to Len(toConv)
		   tempChar = Mid(toConv, i, 1)
		  stringToByte = stringToByte & chrB(AscB(tempChar))
		 Next
	  End Function

	  Private Function byteToString(toConv)
		For i = 1 to LenB(toConv)
		  byteToString = byteToString & chr(AscB(MidB(toConv,i,1))) 
		Next
	  End Function

	class devedit
	
		private e__controlName
		private e__controlWidth
		private e__controlHeight
		private e__initialValue
		private e__initialValueNoBase
		private e__langPack
		private e__hideSave
		private e__hideSpelling
		private e__hideRemoveTextFormatting
		private e__hideFullScreen
		private e__hideBold
		private e__hideUnderline
		private e__hideItalic
		private e__hideStrikethrough
		private e__hideNumberList
		private e__hideBulletList
		private e__hideDecreaseIndent
		private e__hideIncreaseIndent
		private e__hideSuperScript
		private e__hideSubScript
		private e__hideLeftAlign
		private e__hideCenterAlign
		private e__hideRightAlign
		private e__hideJustify
		private e__hideHorizontalRule
		private e__hideLink
		private e__hideAnchor
		private e__hideMailLink
		private e__hideHelp
		private e__hideFont
		private e__hideSize
		private e__hideFormat
		private e__hideStyle
		private e__hideForeColor
		private e__hideBackColor
		private e__hideTable
		private e__hideForm
		private e__hideImage
		private e__hideFlash
		private e__flashPath
		private e__hideTextBox
		private e__hideSymbols
		private e__hideProps
		private e__hideClean
		private e__hideWord
		private e__hideAbsolute
		private e__hideGuidelines
		private e__disableSourceMode
		private e__disablePreviewMode
		private e__guidelinesOnByDefault
		private e__imagePathType
		private e__docType
		private e__imageDisplayType
		private e__flashDisplayType
		private e__disableImageUploading
		private e__disableImageDeleting
		private e__disableFlashUploading
		private e__disableFlashDeleting
		private e__enableXHTMLSupport
		private e__useSingleLineReturn
		private e__customInsertArray
		private e__customLinkArray
		private e__hasCustomInserts
		private e__hasCustomLinks
		private e__snippetCSS
		private e__textareaCols
		private e__textareaRows
		private e__fontNameList
		private e__fontSizeList
		private e__hideWebImage
		private e__hideWebFlash
		private e__language
		private e__imageLibsArray
		private e__hasImageLibraries
		private e__imagePath
		private e__hasFlashLibraries
		private e__flashLibsArray
		private e__baseHref
		
		'Keep track of how many buttons are hidden in the top row.
		'If they are all hidden, then we dont show that row of the menu.
		private e__numTopHidden
		private e__numBottomHidden

		public sub Class_Initialize()

			'Set the default value of all private variables for the class
			 e__controlName = ""
			 e__controlWidth = 0
			 e__controlHeight = 0
			 e__initialValue = ""
			 e__initialValueNoBase = ""
			 e__langPack = 0
			 e__hideSave = 0
			 e__hideSpelling = 0
			 e__hideRemoveTextFormatting = 0
			 e__hideFullScreen = 0
			 e__hideBold = 0
			 e__hideUnderline = 0
			 e__hideItalic = 0
			 e__hideStrikethrough = 0
			 e__hideNumberList = 0
			 e__hideBulletList = 0
			 e__hideDecreaseIndent = 0
			 e__hideIncreaseIndent = 0
			 e__hideSuperScript = 0
			 e__hideSubScript = 0
			 e__hideLeftAlign = 0
			 e__hideCenterAlign = 0
			 e__hideRightAlign = 0
			 e__hideJustify = 0
			 e__hideHorizontalRule = 0
			 e__hideLink = 0
			 e__hideAnchor = 0
			 e__hideMailLink = 0
			 e__hideHelp = 0
			 e__hideFont = 0
			 e__hideSize = 0
			 e__hideFormat = 0
			 e__hideStyle = 0
			 e__hideForeColor = 0
			 e__hideBackColor = 0
			 e__hideTable = 0
			 e__hideForm = 0
			 e__hideImage = 0
			 e__hideFlash = 0
			 e__flashPath = ""
			 e__hideTextBox = 0
			 e__hideSymbols = 0
			 e__hideProps = 0
			 e__hideWord = 0
			 e__hideClean = 0
			 e__hideAbsolute = 0
			 e__hideGuidelines = 0
			 e__disableSourceMode = 0
			 e__disablePreviewMode = 0
			 e__guidelinesOnByDefault = 0
 			 e__numTopHidden = 0
			 e__numBottomHidden = 0
			 e__imagePathType = 0
			 e__docType = 0
			 e__imageDisplayType = 0
			 e__flashDisplayType = 0
			 e__disableImageUploading = 0
			 e__disableImageDeleting = 0
			 e__disableFlashUploading = 0
			 e__disableFlashDeleting = 0
			 e__enableXHTMLSupport = 1
			 e__useSingleLineReturn = 1
			 set e__customInsertArray = Server.CreateObject("Scripting.Dictionary")
			 set e__customLinkArray = Server.CreateObject("Scripting.Dictionary")
			 e__hasCustomInserts = false
			 e__hasCustomLinks = false
			 e__snippetCSS = ""
			 e__textareaCols = 30
			 e__textareaRows = 10
			 e__fontNameList = array()
			 e__fontSizeList = array()
			 e__hideWebImage = 0
			 e__hideWebFlash = 0
			 e__language = de_AMERICAN
			 set e__imageLibsArray = Server.CreateObject("Scripting.Dictionary")
			 e__hasImageLibraries = false
			 e__imagePath = ""
			 set e__flashLibsArray = Server.CreateObject("Scripting.Dictionary")
			 e__hasFlashLibraries = false
			 e__baseHref = ""

		end sub

		public sub SetName(CtrlName)
			e__controlName = CtrlName
		end sub

		public sub SetWidth(Width)
			e__controlWidth = Width
		end sub
		
		public sub SetHeight(Height)
			e__controlHeight = Height
		end sub

		' modified for v5.0
		public sub SetBaseHref(BaseHref)
			e__baseHref = BaseHref
		end sub

		public sub SetFlashPath(FlashPath)
			e__flashPath = FlashPath
		end sub
		
		public sub SetValue(HTMLValue)

			' modified for v5.0
			e__initialValueNoBase = HTMLValue

			if e__docType = de_DOC_TYPE_SNIPPET then
				if e__baseHref <> "" then
					HTMLValue = "<base href=" & e__baseHref & "><body>" & HTMLValue & "</body>"
				else
					HTMLValue = "<body>" & HTMLValue & "</body>"
					e__initialValueNoBase = "<body>" & e__initialValueNoBase & "</body>"
				end if
			end if

			if e__docType = de_DOC_TYPE_SNIPPET and e__snippetCSS <> "" then
				HTMLValue = "<link rel='stylesheet' type='text/css' href='" & e__snippetCSS & "'>" & HTMLValue
				e__initialValueNoBase = "<link rel='stylesheet' type='text/css' href='" & e__snippetCSS & "'>" & e__initialValueNoBase
			end if

			if e__docType = de_DOC_TYPE_HTML_PAGE and instr(lcase(HTMLValue), "<body") = 0 then
				HTMLValue = "<body>" & HTMLValue & "</body>"
				e__initialValueNoBase = "<body>" & e__initialValueNoBase & "</body>"
			end if

			' end modification
				
			'Format the initial text so that we can set the content of the iFrame to its value
			e__initialValue = HTMLValue

			if e__initialValue <> "" then

				if isIE55OrAbove = true then
					e__initialValue = HTMLValue
					e__initialValue = replace(e__initialValue, "\", "\\")
       				e__initialValue = replace(e__initialValue, "'", "\'")
					e__initialValue = replace(e__initialValue, "&#39;", "\&#39;")
       				e__initialValue = replace(e__initialValue, chr(13), "")
       				e__initialValue = replace(e__initialValue, chr(10), "")

					e__initialValueNoBase = replace(e__initialValueNoBase, "\", "\\")
       				e__initialValueNoBase = replace(e__initialValueNoBase, "'", "\'")
					e__initialValueNoBase = replace(e__initialValueNoBase, "&#39;", "\&#39;")
       				e__initialValueNoBase = replace(e__initialValueNoBase, chr(13), "")
       				e__initialValueNoBase = replace(e__initialValueNoBase, chr(10), "")
				else
					e__initialValue = HTMLValue
				end if

			end if

		end sub

		public function GetValue(ConvertQuotes)
		
			dim tmpVal
			
			tmpVal = Request.Form(e__controlName & "_html")

			if ConvertQuotes = true then
				tmpVal = Replace(tmpVal, "'", "''")
			end if
			
			GetValue = tmpVal
		
		end function

		public sub HideSaveButton()

			'Hide the save button
			e__hideSave = true
			e__numTopHidden = e__numTopHidden + 1

		end sub

		public sub HideSpellingButton()

			' Hide the spelling button
			e__hideSpelling = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideRemoveTextFormattingButton()

			' Hide the remove text formatting button
			e__hideRemoveTextFormatting = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideFullScreenButton()
		
			'Hide the full screen button
			e__hideFullScreen = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideBoldButton()
		
			'Hide the bold button
			e__hideBold = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideUnderlineButton()
		
			'Hide the underline button
			e__hideUnderline = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideItalicButton()
		
			'Hide the italic button
			e__hideItalic = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideStrikethroughButton()

			'Hide the strikethrough button
			e__hideStrikethrough = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideNumberListButton()
		
			'Hide the number list button
			e__hideNumberList = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideBulletListButton()
		
			'Hide the bullet list button
			e__hideBulletList = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideDecreaseIndentButton()
		
			'Hide the decrease indent button
			e__hideDecreaseIndent = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideIncreaseIndentButton()
		
			'Hide the increase indent button
			e__hideIncreaseIndent = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideSuperScriptButton()
		
			'Hide the superscript button
			e__hideSuperScript = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideSubScriptButton()
		
			'Hide the subscript button
			e__hideSubScript = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideLeftAlignButton()
		
			'Hide the left align button
			e__hideLeftAlign = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideCenterAlignButton()
		
			'Hide the center align button
			e__hideCenterAlign = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideRightAlignButton()
		
			'Hide the right align button
			e__hideRightAlign = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideJustifyButton()
		
			'Hide the left align button
			e__hideJustify = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideHorizontalRuleButton()
		
			'Hide the horizontal rule button
			e__hideHorizontalRule = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideLinkButton()
		
			'Hide the link button
			e__hideLink = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideAnchorButton()
		
			'Hide the anchor button
			e__hideAnchor = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideMailLinkButton()
		
			'Hide the mail link button
			e__hideMailLink = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub

		public sub HideHelpButton()
		
			'Hide the help button
			e__hideHelp = true
			e__numTopHidden = e__numTopHidden + 1
		
		end sub
		
		public sub HideFontList()
		
			'Hide the font list
			e__hideFont = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideSizeList()
		
			'Hide the size list
			e__hideSize = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub HideFormatList()
		
			'Hide the format list
			e__hideFormat = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideStyleList()
		
			'Hide the style list
			e__hideStyle = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideForeColorButton()
		
			'Hide the forecolor button
			e__hideForeColor = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub HideBackColorButton()
		
			'Hide the backcolor button
			e__hideBackColor = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub HideTableButton()
		
			'Hide the table button
			e__hideTable = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub HideFormButton()
		
			'Hide the form button
			e__hideForm = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideImageButton()
		
			'Hide the image button
			e__hideImage = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideFlashButton()
		
			'Hide the flash button
			e__hideFlash = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideTextBoxButton()

			'Hide the textbox button
			e__hideTextBox = true
			e__numBottomHidden = e__numBottomHidden + 1

		end sub

		public sub HideSymbolButton()
		
			'Hide the symbol button
			e__hideSymbols = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HidePropertiesButton()
		
			'Hide the properties button
			e__hideProps = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideCleanHTMLButton()
		
			'Hide the clean HTML button
			e__hideClean = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideAbsolutePositionButton()
		
			'Hide the absolute position button
			e__hideAbsolute = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub

		public sub HideGuidelinesButton()
		
			'Hide the guidelines button
			e__hideGuidelines = true
			e__numBottomHidden = e__numBottomHidden + 1
		
		end sub
		
		public sub DisableSourceMode()
		
			'Hide the source mode button
			e__disableSourceMode = true
		
		end sub
		
		public sub DisablePreviewMode()
		
			'Hide the preview mode button
			e__disablePreviewMode = true
		
		end sub
		
		public sub EnableGuidelines()

			'Set the table guidelines on by default
			e__guidelinesOnByDefault = 1

		end sub

		public sub SetPathType(PathType)

			'How do we want to include the path to the images? 0 = Full, 1 = Absolute
			e__imagePathType = PathType

		end sub

		public sub SetDocumentType(DocType)

			'Is the user editing a full HTML document
			e__docType = DocType

		end sub

		public sub SetImageDisplayType(DisplayType)

			'How should the images be displayed in the image manager? 0 = Line / 1 = Thumbnails
			e__imageDisplayType = DisplayType
		
		end sub

		public sub DisableImageUploading()

			'Do we need to stop images being uploaded?
			e__disableImageUploading = 1

		end sub

		public sub DisableImageDeleting()

			'Do we need to stop images from being delete?
			e__disableImageDeleting = 1
		
		end sub

		public sub SetFlashDisplayType(DisplayType)

			'How should the flash files be displayed in the image manager? 0 = Line / 1 = Thumbnails
			e__flashDisplayType = DisplayType
		
		end sub

		public sub DisableFlashUploading()

			'Do we need to stop flash files being uploaded?
			e__disableFlashUploading = 1
		
		end sub

		public sub DisableFlashDeleting()

			'Do we need to stop flash files from being deleted?
			e__disableFlashDeleting = 1
		
		end sub

		public function isIE55OrAbove()

			' Is it MSIE?
			dim browserCheck1, browserCheck2, browserCheck3, browserCheck4, browserCheck5, browserCheck6

			browserCheck1 = instr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE")

			if browserCheck1 > 0 then
				browserCheck1 = true
			else
				browserCheck1 = false
			end if

			' Is it NOT Opera?
			browserCheck2 = instr(1, Request.ServerVariables("HTTP_USER_AGENT"), "Opera")

			if browserCheck2 = 0 then
				browserCheck2 = true
			else
				browserCheck2 = false
			end if

			' Is it MSIE 5 or above?
			browserCheck3 = instr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE 5.5")
			browserCheck4 = instr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE 6")

			if browserCheck3 > 0 OR browserCheck4 > 0 then
				browserCheck5 = true
			else
				browserCheck5 = false
			end if

			browserCheck6 = instr(1, Request.ServerVariables("HTTP_USER_AGENT"), "Windows")

			if browserCheck6 = 0 then
				browserCheck6 = false
			else
				browserCheck6 = true
			end if

			if browserCheck1 = true AND browserCheck2 = true AND browserCheck5 = true AND browserCheck6 = true then
				isIE55OrAbove = true
			else
				isIE55OrAbove = false
			end if

		end function

		' -------------------------
		' Version 3.0 new functions

		public function DisableXHTMLFormatting()

			' Disable XHTML formatting of inline code
			e__enableXHTMLSupport = 0
		
		end function
		
		public function DisableSingleLineReturn()
		
			' Instead of adding a <p> tag for a new line, add <br> instead
			e__useSingleLineReturn = 0
		
		end function

		public function LoadHTMLFromAccessQuery(ByVal DatabaseFile, ByVal DatabaseQuery, ByRef ErrorDesc)
		
			' Grabs a value from an Access database based on a SELECT query.
			' It will return a text value from the field on success, or false on failure
			
			Err.Clear
			On Error Resume Next
			
			dim aConn
			dim aRS
			
			Set aConn =  Server.CreateObject("ADODB.Connection") 
			aConn.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & DatabaseFile & ";"
			
			if Err.number <> 0 then
				ErrorDesc = Err.Description
				LoadHTMLFromAccessQuery = false
				exit function
			else
				Set aRS = aConn.Execute(DatabaseQuery)
				
				if Err.number <> 0 then
					ErrorDesc = Err.Description
					LoadHTMLFromAccessQuery = false
					exit function
				else
					if aRS.EOF then
						ErrorDesc = Err.Description
						LoadHTMLFromAccessQuery = false
						exit function
					else
						'We have a field value, set it and return
						SetValue aRS.Fields(0).value
						LoadHTMLFromAccessQuery = true
						exit function
					end if
				end if
			end if
			
		end function

		public function LoadHTMLFromSQLServerQuery(ByVal DatabaseServer, ByVal DatabaseName, ByVal DatabaseUser, ByVal DatabasePassword, ByVal DatabaseQuery, ByRef ErrorDesc)
		
			' Grabs a value from an SQL Server database based on a SELECT query.
			' It will return a text value from the field on success, or false on failure
			
			'Err.Clear
			'On Error Resume Next
			
			dim aConn
			dim aRS
			
			Set aConn =  Server.CreateObject("ADODB.Connection") 
			aConn.Open "Provider=SQLOLEDB; Data Source = " & DatabaseServer & "; Initial Catalog = " & DatabaseName & "; User Id = " & DatabaseUser & "; Password=" & DatabasePassword
			
			if Err.number <> 0 then
				ErrorDesc = Err.Description
				LoadHTMLFromSQLServerQuery = false
				exit function
			else
				Set aRS = aConn.Execute(DatabaseQuery)
				
				if Err.number <> 0 then
					ErrorDesc = Err.Description
					LoadHTMLFromSQLServerQuery = false
					exit function
				else
					if aRS.EOF then
						ErrorDesc = Err.Description
						LoadHTMLFromSQLServerQuery = false
						exit function
					else
						'We have a field value, set it and return
						SetValue aRS.Fields(0).value
						LoadHTMLFromSQLServerQuery = true
						exit function
					end if
				end if
			end if
			
		end function
		
		public function LoadFromFile(ByVal FilePath, ByRef ErrorDesc)
		
			' This function will load the entire text from a file and
			' set the value of the DevEdit control
			
			Err.Clear
			On Error Resume Next
			
			dim fso
			dim file
			dim fileContents
			
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			FilePath = Server.mapPath(FilePath)
			
			if fso.FileExists(FilePath) then
				'File exists, open it and read it in
				set file = fso.OpenTextFile(FilePath, 1, false, -2)
				
				if err.number <> 0 then
					'An error occured while opening the file
					ErrorDesc = Err.Description
					LoadFromFile = false
					exit function
				else
					'The file was loaded OK, read it into a variable
					fileContents = file.ReadAll
					SetValue fileContents
					LoadFromFile = true
					exit function
				end if
			else
				'File doesnt exist
				ErrorDesc = "File " & FilePath & " doesn't exist"
				LoadFromFile = false
				exit function
			end if

		end function

		public function SaveToFile(ByVal FilePath, ByRef ErrorDesc)
		
			' Writes the contents of the DevEdit control to a file
			Err.Clear
			On Error Resume Next
			
			dim fso
			dim file
			dim stream
			dim fileContents
			
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			FilePath = Server.MapPath(FilePath)
			
			if len(GetValue(false)) = 0 then
				' No data to write to the file
				ErrorDesc = "Cannot save an empty value to " & FilePath
				SaveToFile = false
				exit function
			else
				if NOT fso.FileExists(FilePath) then
					'Attempt to create the file
					set file = fso.CreateTextFile(FilePath)
					
					if err.number <> 0 then
						ErrorDesc = err.Description
						SaveToFile = false
						exit function
					end if
				end if
			
				' Now that the file definetly exists, we can attempt to open it for writing
				set file = fso.GetFile(FilePath)
			
				if err.number <> 0 then
					'Failed to grab the file
					ErrorDesc = err.Description
					SaveToFile = false
					exit function
				else
					'The file was opened OK, write to it
					set stream = file.OpenAsTextStream(2)
					
					if err.number <> 0 then
						'Failed to open the file for writing
						ErrorDesc = err.Description
						SaveToFile = false
						exit function
					else
						'Opened the file for writing OK
						stream.Write(GetValue(false))
						stream.Close
						SaveToFile = true
						exit function
					end if
				end if
			end if
		
		end function

		public function AddCustomInsert(InsertName, InsertHTMLCode)
		
			e__hasCustomInserts = true
			
			if e__customInsertArray.Exists(InsertName) = false then
				e__customInsertArray.Add InsertName, InsertHTMLCode
			end if

		end function

		'Added for v5.0
		public function AddCustomLink(LinkName, LinkURL, TargetWindow)
		
			e__hasCustomLinks = true
			
			if e__customLinkArray.Exists(LinkName) = false then
				e__customLinkArray.Add LinkName, LinkURL & "|" & TargetWindow
			end if

		end function

		public function AddImageLibrary(LibraryName, LibraryPath)

			e__hasImageLibraries = true

			if e__imageLibsArray.Exists(LibraryName) = false then
				e__imageLibsArray.Add LibraryName, LibraryPath
			end if

		end function

		public function AddFlashLibrary(LibraryName, LibraryPath)

			e__hasFlashLibraries = true

			if e__flashLibsArray.Exists(LibraryName) = false then
				e__flashLibsArray.Add LibraryName, LibraryPath
			end if

		end function

		private function FormatCustomInsertText()
		
			' Private Function - This function will return all of the custom inserts as JavaScript arrays
			dim i, name, html, ciText, keys, vals
			
			if e__hasCustomInserts = true then
				
				ciText = "["
				keys = e__customInsertArray.Keys
				vals = e__customInsertArray.Items
			
				for i = 0 to e__customInsertArray.Count-1
					name = replace(replace(keys(i), """", "\"""), chr(13) & chr(10), "\r\n")
					html = replace(replace(vals(i), """", "\"""), chr(13) & chr(10), "\r\n")
					ciText = ciText & "[""" & name & """, """ & html & """],"
				next
				
				ciText = left(ciText, len(ciText)-1)
				ciText = ciText & "]"
				
			else
				ciText = "[]"
			end if
			
			FormatCustomInsertText = ciText
		
		end function

		private function FormatCustomLinkText()
		
			' Private Function - This function will return all of the custom links as JavaScript arrays
			dim i, name, url, target, ciText, keys, vals
			
			if e__hasCustomLinks = true then
				
				ciText = "["
				keys = e__customLinkArray.Keys
				vals = e__customLinkArray.Items
			
				for i = 0 to e__customLinkArray.Count-1
					name = replace(replace(keys(i), """", "\"""), chr(13) & chr(10), "\r\n")
					url = replace(replace(vals(i), """", "\"""), chr(13) & chr(10), "\r\n")
					ciText = ciText & "[""" & name & """, """ & url & """],"
				next
				
				ciText = left(ciText, len(ciText)-1)
				ciText = ciText & "]"
				
			else
				ciText = "[]"
			end if
			
			FormatCustomLinkText = ciText
		
		end function

		public function GetImageLibraries()

			' Private Function - This function will return all of the image libraries as <SELECT> options
			dim i, name, dir, ciText, keys, vals

			ciText = "<option value=""" & e__imagePath & """>" & sTxtDefaultImageLibrary & "</option>"

			if e__hasImageLibraries = true then

				keys = e__imageLibsArray.Keys
				vals = e__imageLibsArray.Items

				for i = 0 to e__imageLibsArray.Count-1
					name = replace(replace(keys(i), """", "\"""), chr(13) & chr(10), "\r\n")
					dir = replace(replace(vals(i), """", "\"""), chr(13) & chr(10), "\r\n")
					ciText = ciText & "<option value=""" & dir & """>" & name & "</option>"
				next
				
			end if

			GetImageLibraries = replace(ciText, "'", "\'")
		
		end function

		public function GetFlashLibraries()

			' Private Function - This function will return all of the flash libraries as <SELECT> options
			dim i, name, dir, ciText, keys, vals

			ciText = "<option value=""" & e__flashPath & """>" & sTxtDefaultFlashLibrary & "</option>"

			if e__hasFlashLibraries = true then

				keys = e__flashLibsArray.Keys
				vals = e__flashLibsArray.Items

				for i = 0 to e__flashLibsArray.Count-1
					name = replace(replace(keys(i), """", "\"""), chr(13) & chr(10), "\r\n")
					dir = replace(replace(vals(i), """", "\"""), chr(13) & chr(10), "\r\n")
					ciText = ciText & "<option value=""" & dir & """>" & name & "</option>"
				next
				
			end if

			GetFlashLibraries = replace(ciText, "'", "\'")
		
		end function

		public function SetSnippetStyleSheet(StyleSheetURL)

			' Sets the location of the stylesheet for a code snippet
			e__docType = de_DOC_TYPE_SNIPPET
			e__snippetCSS = StyleSheetURL

		end function
		
		public function SetTextAreaDimensions(Cols, Rows)
		
			' Sets the rows and cols attributes of the <textarea> tag that will appear
			' if the client isnt using Internet explorer
			e__textareaCols = Cols
			e__textareaRows = Rows
		
		end function

		' End Version 3.0 new functions
		' Version 4.0 new functions

		public sub SetLanguage(Lang)

			select case(CStr(Lang))
				case "1"
					e__language = "american"
				case "2"
					e__language = "british"
				case "3"
					e__language = "canadian"
				case "4"
					e__language = "french"
				case "5"
					e__language = "spanish"
				case "6"
					e__language = "german"
				case "7"
					e__language = "italian"
				case "8"
					e__language = "portugese"
				case "9"
					e__language = "dutch"
				case "10"
					e__language = "norwegian"
				case "11"
					e__language = "swedish"
				case "12"
					e__language = "danish"
				case else
					e__language = "american"
			end select

		end sub

		public function DisableInsertImageFromWeb

			e__hideWebImage = 1

		end function

		'modified for v5.0
		public function DisableInsertFlashFromWeb

			e__hideWebFlash = 1

		end function

		public function BuildSizeList()

			%><option selected><%=sTxtSize%></option><%

			if uBound(e__fontSizeList) >= 0 then

				' Build the list of font sizes from the list that the user has specified
				for i = 0 to uBound(e__fontSizeList)
					%><option value="<%=trim(e__fontSizeList(i)) %>"><%=trim(e__fontSizeList(i)) %></option><%
				next

			else

				' Build the list of font sizes manually
				%>
					<option value="1">1</option>
			  		<option value="2">2</option>
			  		<option value="3">3</option>
			  		<option value="4">4</option>
			  		<option value="5">5</option>
			  		<option value="6">6</option>
			  		<option value="7">7</option>
				<%
			end if

		end function

		public function BuildFontList()

			%><option selected><%=sTxtFont%></option><%

			if uBound(e__fontNameList) >= 0 then

				' Build the list of fonts from the list that the user has specified
				for i = 0 to uBound(e__fontNameList)
					%><option value="<%=trim(e__fontNameList(i)) %>"><%=trim(e__fontNameList(i)) %></option><%
				next

			else

				' Build the list of fonts manually
				%>
					<option value="Times New Roman">Default</option>
					<option value="Arial">Arial</option>
					<option value="Verdana">Verdana</option>
					<option value="Tahoma">Tahoma</option>
					<option value="Courier New">Courier New</option>
					<option value="Georgia">Georgia</option>
				<%
			end if

		end function

		public function SetFontList(FontList)

			dim tmpFontList
			tmpFontList = split(FontList, ",")

			if isArray(tmpFontList) then
				e__fontNameList = tmpFontList
			end if

		end function

		public function SetFontSizeList(SizeList)

			dim tmpSizeList
			tmpSizeList = split(SizeList, ",")

			if isArray(tmpSizeList) then
				e__fontSizeList = tmpSizeList
			end if

		end function

		' End Version 4.0 new functions
		' -------------------------

		public sub ShowControl(Width, Height, ImagePath)

			SetWidth(Width)
			SetHeight(Height)
			e__imagePath = ImagePath

			if e__controlName = "" then
				Response.Write "<b>ERROR: Must set an DevEdit control name using the SetName() function</b>"
				Response.End
			end if

			if DevEditPath = "" then
				Response.Write "<b>ERROR: Must set the path to the DevEdit folder using the SetDevEditPath() function</b>"
				Response.End
			end if

			' If the browser isn't IE5.5 or above, show a <textarea> tag and die
			if isIE55OrAbove = false then
			%>
				<span style="background-color: lightyellow"><font face="verdana" size="1" color="red"><b><%=sTxtTextAreaError%></b></font></span><br>
				<textarea style="width:<%=e__controlWidth %>; height:<%=e__controlHeight%>" rows="<%=e__textareaRows%>" cols="<%=e__textareaCols%>" name="<%=e__controlName %>_html"><%=e__initialValue%></textarea>
			<%
			else

					'Output the hidden textarea buffer tag which will contain the iFrame source
					Response.Write "<textarea style=display:none id='" & e__controlName & "_src'>"

					%><link rel="stylesheet" href="<%=DevEditPath%>/de_includes/de_styles.css" type="text/css"><%

        			'Do we need to hide the page properties button?
        			if e__hideProps <> 0 or e__docType = 0 then
        				HidePropertiesButton
        			end if
        			
        			dim URL, scriptName, scriptDir, slashPos

					URL = Request.ServerVariables("http_host")
					scriptName = DevEditPath & "/class.devedit.asp"
					
					'Workout the location of class.devedit.asp
					scriptDir = strreverse(Request.ServerVariables("path_info"))
					slashPos = instr(1, scriptDir, "/")
					scriptDir = strreverse(mid(scriptDir, slashPos, len(scriptDir)))
					scriptName = scriptDir & scriptName

					if e__enableXHTMLSupport = 1 then
					' modified for v5.0
					%>
						<script language="JavaScript" src="<%=DevEditPath%>/de_includes/de_xhtml_define.js" type="text/javascript"></script>
						<script language="JavaScript" src="<%=DevEditPath%>/de_includes/de_xhtml.js" type="text/javascript"></script>
					<%
					' end mod
					end if
					%>
					<script>
							var customInserts = <%=FormatCustomInsertText%>;
							var customLinks = <%=FormatCustomLinkText%>;
							var tableDefault = <%=e__guidelinesOnByDefault%>;
							var useBR = <%=e__useSingleLineReturn%>;
							var useXHTML = "<%=e__enableXHTMLSupport%>";
							var ContextMenuWidth = <%=sTxtContextMenuWidth%>;
							var URL = "<%=URL%>";
							var ScriptName = "<%=DevEditPath%>/class.devedit.asp";
							var sTxtGuidelines = "<%=sTxtGuidelines%>";
							var sTxtOn = "<%=sTxtOn%>";
							var sTxtOff = "<%=sTxtOff%>";
							var sTxtClean = "<%=sTxtClean%>";
							// var re2 = /href="<%=HTTPStr%>:\/\/<%=URL%>/g
							var re3 = /src="<%=HTTPStr%>:\/\/<%=URL%>/g;
							var re4 = /src="<%=HTTPStr%>:\/\/<%=URL%>/g;
							var re5 = /src="http:\/\/<%=URL%>/g;
							var isEditingHTMLPage = <%=e__docType%>;
							var pathType = <%=e__imagePathType%>;
							var imageDir = "<%=Server.URLEncode(ImagePath)%>";
							var flashDir = "<%=Server.URLEncode(e__flashPath)%>";
							var showThumbnails = <%=e__imageDisplayType%>;
							var showFlashThumbnails = <%=e__flashDisplayType%>;
							var disableImageUploading = <%=e__disableImageUploading%>;
							var disableImageDeleting = <%=e__disableImageDeleting%>;
							var disableFlashUploading = <%=e__disableFlashUploading%>;
							var disableFlashDeleting = <%=e__disableFlashDeleting%>;
							var HideWebImage = <%=e__hideWebImage%>;
							var HideWebFlash = <%=e__hideWebFlash%>;
							var HTTPStr = "<%=HTTPStr%>";
							var spellLang = "<%=e__language%>";
							var controlName = "<%=e__controlName%>_frame";
							var imageLibs = '<%=GetImageLibraries%>';
							var flashLibs = '<%=GetFlashLibraries%>';
							var myBaseHref = '<%=e__baseHref%>';
							var deveditPath = '<%=DevEditPath%>';
							var deveditPath1 = '<%=DevEditPath%>';
					</script>

					<script>
						<% DisplayIncludes "enc_functions.js", "Javascript Functions" %>
					</script>
					<script language="JavaScript" src="<%=DevEditPath%>/de_includes/de_functions.js" type="text/javascript"></script>

					<!-- modified for v5.0 -->
					<span id="fooContainer" style="width:100%; border:1px ridge">
        			<%
        
        			'---------------------------------------
					'                                      '
					' DevEdit Toolbar                      '
					'                                      '
					'--------------------------------------'
					%>
						<table width="100%" cellspacing="0" cellpadding="0" class=toolbar>
							<tr>
							<td class="body" height="24">
							<table width="100%" border="0" cellspacing="0" cellpadding="0" class="hide" align="center" id="toolbar_preview">
								<tr>
								  <td class="body" height="24">
								  &nbsp;<img src="<%=DevEditPath%>/de_images/popups/preview.gif" width="21" height="20" align=absmiddle>&nbsp;Preview Mode
								  </td>
								 </tr>
							</table>
							 <table width="100%" border="0" cellspacing="0" cellpadding="0" class="hide" align="center" id="toolbar_code">
								<tr>
								  <td class="body" height="22">
								  <table border="0" cellspacing="0" cellpadding="1">
									  <tr id=de>
										<% if e__hideSave <> true then %>
											<td>
												<img border="0" src="<%=DevEditPath%>/de_images/button_save.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='saveDevEdit();' title="<%=sTxtSave%>" class=toolbutton></td>
										<% end if %>
										<% if e__hideFullScreen <> true then %>
											<td>
												<img id=fullscreen2 border="0" src="<%=DevEditPath%>/de_images/button_fullscreen.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='toggleSize();foo.focus();' title="<%=sTxtFullscreen%>" class=toolbutton></td>
										<% end if %>
										<td>
										  <img border="0" disabled id="toolbarCut2_off" src="<%=DevEditPath%>/de_images/button_cut_disabled.gif" width="21" height="20" title="<%=sTxtCut%> (Ctrl+X)" class=toolbutton><img border="0" id="toolbarCut2_on" src="<%=DevEditPath%>/de_images/button_cut.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Cut");foo.focus();' title="<%=sTxtCut%> (Ctrl+X)" class=toolbutton style="display:none"></td>
										<td>
										  <img border="0" disabled id="toolbarCopy2_off" src="<%=DevEditPath%>/de_images/button_copy_disabled.gif" width="21" height="20" title="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton><img border="0" id="toolbarCopy2_on" src="<%=DevEditPath%>/de_images/button_copy.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Copy");foo.focus();' title="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton style="display:none"></td>
										<td>
										  <img border="0" disabled id="toolbarPasteButton2_off" src="<%=DevEditPath%>/de_images/button_paste_disabled.gif" width="21" height="20" title="<%=sTxtPaste%> (Ctrl+V)" class=toolbutton><img border="0" id="toolbarPasteButton2_on" src="<%=DevEditPath%>/de_images/button_paste.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Paste");foo.focus();' title="<%=sTxtPaste%> (Ctrl+V)" class=toolbutton style="display:none"></td>
										<td>
										  <img border="0" src="<%=DevEditPath%>/de_images/button_find.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='ShowFindDialog();foo.focus();' title="<%=sTxtFindReplace%>" class=toolbutton></td>
										<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										<td>
										  <img border="0" disabled id="undo2_off" src="<%=DevEditPath%>/de_images/button_undo_disabled.gif" width="21" height="20" title="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton><img border="0" id="undo2_on" src="<%=DevEditPath%>/de_images/button_undo.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Undo");' title="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton style="display:none"></td>
										<td>
										  <img border="0" disabled id="redo2_off" src="<%=DevEditPath%>/de_images/button_redo_disabled.gif" width="21" height="20" title="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton><img border="0" id="redo2_on" src="<%=DevEditPath%>/de_images/button_redo.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Redo");' title="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton style="display:none"></td>
										</tr>
									</table>
								  </td>
								 </tr>
								<tr>
								  <td class="body" bgcolor="#808080"><img src="<%=DevEditPath%>/de_images/1x1.gif" width="1" height="1"></td>
								</tr>
								<tr>
								  <td class="body" bgcolor="#FFFFFF"><img src="<%=DevEditPath%>/de_images/1x1.gif" width="1" height="1"></td>
								</tr>
								 <tr><td height=24>&nbsp;</td></tr>
							</table>
							  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="bevel3" align="center" id="toolbar_full">
								<tr>
								  <td class="body" height="22">
									<table border="0" cellspacing="0" cellpadding="1" id=toolbar1>
									  <tr id=de>
										  <% if e__hideSave <> true then %>
											<td>
												<img border="0" src="<%=DevEditPath%>/de_images/button_save.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='saveDevEdit();' title="Save" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideFullScreen <> true then %>
											<td>
												<img id=fullscreen border="0" src="<%=DevEditPath%>/de_images/button_fullscreen.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='toggleSize();foo.focus();' title="<%=sTxtFullscreen%>" class=toolbutton></td>
										  <% end if %>
											<td>
												<img border="0" disabled id="toolbarCut_off" src="<%=DevEditPath%>/de_images/button_cut_disabled.gif" width="21" height="20" title="<%=sTxtCut%> (Ctrl+X)" class=toolbutton><img border="0" id="toolbarCut_on" src="<%=DevEditPath%>/de_images/button_cut.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Cut");foo.focus();' title="<%=sTxtCut%> (Ctrl+X)" class=toolbutton style="display:none"></td>
											<td>
												<img border="0" disabled id="toolbarCopy_off" src="<%=DevEditPath%>/de_images/button_copy_disabled.gif" width="21" height="20" title="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton><img border="0" id="toolbarCopy_on" src="<%=DevEditPath%>/de_images/button_copy.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Copy");foo.focus();' title="<%=sTxtCopy%> (Ctrl+C)" class=toolbutton style="display:none"></td>
											<td>
												<img id=toolbarPasteButton_off disabled class=toolbutton width="21" height="20" src="<%=DevEditPath%>/de_images/button_paste_disabled.gif" border=0 unselectable="on" title="<%=sTxtPaste%> (Ctrl+V)"><img id=toolbarPasteButton_on class=toolbutton onMouseDown="button_down(this);" onMouseOver="button_over(this); button_over(toolbarPasteDrop_on)" onClick="doCommand('Paste'); foo.focus()" onMouseOut="button_out(this); button_out(toolbarPasteDrop_on);" width="21" height="20" src="<%=DevEditPath%>/de_images/button_paste.gif" border=0 unselectable="on" title="<%=sTxtPaste%> (Ctrl+V)" style="display:none"><img id=toolbarPasteDrop_off disabled class=toolbutton width="7" height="20" src="<%=DevEditPath%>/de_images/button_drop_menu_disabled.gif" border=0 unselectable="on"><img id=toolbarPasteDrop_on class=toolbutton onMouseDown="button_down(this);" onMouseOver="button_over(this); button_over(toolbarPasteButton_on)" onClick="showMenu('pasteMenu',180,42)" onMouseOut="button_out(this); button_out(toolbarPasteButton_on);" width="7" height="20" src="<%=DevEditPath%>/de_images/button_drop_menu.gif" border=0 unselectable="on" style="display:none"></td>
											<td>
											  <img border="0" src="<%=DevEditPath%>/de_images/button_find.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='ShowFindDialog();foo.focus();' title="<%=sTxtFindReplace%>" class=toolbutton></td>
											<td>
												<img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
											<td>
												<img id="undo_off" disabled UNSELECTABLE="on" border="0" src="<%=DevEditPath%>/de_images/button_undo_disabled.gif" width="21" height="20" title="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton><img id="undo_on" UNSELECTABLE="on" border="0" src="<%=DevEditPath%>/de_images/button_undo.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='goHistory(-1);' title="<%=sTxtUndo%> (Ctrl+Z)" class=toolbutton style="display:none"></td>
											<td>
												<img id="redo_off" disabled UNSELECTABLE="on" border="0" src="<%=DevEditPath%>/de_images/button_redo_disabled.gif" width="21" height="20" title="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton><img id="redo_on" UNSELECTABLE="on" border="0" src="<%=DevEditPath%>/de_images/button_redo.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='goHistory(1);' title="<%=sTxtRedo%> (Ctrl+Y)" class=toolbutton style="display:none"></td>
											<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>

										  <% if e__hideSpelling <> true then %>
											<td>
												<img id="toolbarSpell" border="0" src="<%=DevEditPath%>/de_images/button_spellcheck.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='spellCheck();' title="<%=sTxtCheckSpelling%> (F7)" class=toolbutton></td>
										<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										  <% end if %>

										  <% if e__hideRemoveTextFormatting <> true then %>
											<td>
												<img border="0" src="<%=DevEditPath%>/de_images/button_remove_format.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("RemoveFormat");' title="<%=sTxtRemoveFormatting%>" class=toolbutton></td>
											<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										  <% end if %>
										  <% if e__hideBold <> true then %>
											<td>
												<img id="fontBold" border="0" src="<%=DevEditPath%>/de_images/button_bold.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("Bold");foo.focus();' title="<%=sTxtBold%> (Ctrl+B)" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideUnderline <> true then %>
											<td>
												<img id="fontUnderline" border="0" src="<%=DevEditPath%>/de_images/button_underline.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("Underline");foo.focus();' title="<%=sTxtUnderline%> (Ctrl+U)" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideItalic <> true then %>
											<td>
												<img id="fontItalic" border="0" src="<%=DevEditPath%>/de_images/button_italic.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("Italic");foo.focus();' title="<%=sTxtItalic%> (Ctrl+I)" class=toolbutton></td>
										  <% end if %>

										<% if e__hideStrikethrough <> true then %>
											<td>
												<img id="fontStrikethrough" border="0" src="<%=DevEditPath%>/de_images/button_strikethrough.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("Strikethrough");foo.focus();' title="<%=sTxtStrikethrough%>" class=toolbutton></td>
										<% end if %>
											<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>

										  <% if e__hideNumberList <> true then %>
											<td>
												<img id="fontInsertOrderedList" border="0" src="<%=DevEditPath%>/de_images/button_numbers.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("InsertOrderedList");foo.focus();' title="<%=sTxtNumList%>" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideBulletList <> true then %>
											<td>
												<img id="fontInsertUnorderedList" border="0" src="<%=DevEditPath%>/de_images/button_bullets.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("InsertUnorderedList");foo.focus();' title="<%=sTxtBulletList%>" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideDecreaseIndent <> true then %>
											<td>
											<img border="0" src="<%=DevEditPath%>/de_images/button_decrease_indent.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Outdent");foo.focus();' title="<%=sTxtDecreaseIndent%>" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideIncreaseIndent <> true then %>
											<td>
												<img border="0" src="<%=DevEditPath%>/de_images/button_increase_indent.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("Indent");foo.focus();' title="<%=sTxtIncreaseIndent%>" class=toolbutton></td>
											<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										  <% end if %>
										  <% if e__hideSuperScript <> true then %>
											<td>
												<img id="fontSuperScript" border="0" src="<%=DevEditPath%>/de_images/button_superscript.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("superscript");foo.focus();' title="<%=sTxtSuperscript%>" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideSubScript <> true then %>
											<td>
												<img id="fontSubScript" border="0" src="<%=DevEditPath%>/de_images/button_subscript.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("subscript");foo.focus();' title="<%=sTxtSubscript%>" class=toolbutton></td>
											<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										  <% end if %>
										  <% if e__hideLeftAlign <> true then %>
											<td>
												<img id="fontJustifyLeft" border="0" src="<%=DevEditPath%>/de_images/button_align_left.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("JustifyLeft");foo.focus();' title="<%=sTxtAlignLeft%>" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideCenterAlign <> true then %>
											<td>
												<img id="fontJustifyCenter" border="0" src="<%=DevEditPath%>/de_images/button_align_center.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("JustifyCenter");foo.focus();' title="<%=sTxtAlignCenter%>" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideRightAlign <> true then %>
											<td>
												<img id="fontJustifyRight" border="0" src="<%=DevEditPath%>/de_images/button_align_right.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("JustifyRight");foo.focus();' title="<%=sTxtAlignRight%>" class=toolbutton></td>
										  <% end if %>
										  <% if e__hideJustify <> true then %>
											<td>
												<img id="fontJustifyFull" border="0" src="<%=DevEditPath%>/de_images/button_align_justify.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out2(this);" onmousedown="button_down(this);" onClick='doCommand("JustifyFull");foo.focus();' title="<%=sTxtAlignJustify%>" class=toolbutton></td>
											<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										  <% end if %>
										  <% if e__hideLink <> true then %>
											<td>
												<img disabled id="toolbarLink_off" border="0" src="<%=DevEditPath%>/de_images/button_link_disabled.gif" width="21" height="20" title="<%=sTxtHyperLink%>" class=toolbutton><img id="toolbarLink_on" border="0" src="<%=DevEditPath%>/de_images/button_link.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doLink()' title="<%=sTxtHyperLink%>" class=toolbutton style="display:none"></td>
										  <% end if %>
										  <% if e__hideMailLink <> true then %>
											<td>
												<img border="0" id="toolbarEmail_off" disabled src="<%=DevEditPath%>/de_images/button_email_disabled.gif" width="21" height="20" title="<%=sTxtEmail%>" class=toolbutton><img border="0" id="toolbarEmail_on" src="<%=DevEditPath%>/de_images/button_email.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doEmail()' title="<%=sTxtEmail%>" class=toolbutton style="display:none"></td>
										  <% end if %>
										  <% if e__hideAnchor <> true then %>
											<td>
												<img border="0" src="<%=DevEditPath%>/de_images/button_anchor.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doAnchor()' title="<%=sTxtAnchor%>" class=toolbutton></td>
											<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										  <% end if %>
										  <% if e__hideHelp <> true then %>
											<td>
												<img border="0" src="<%=DevEditPath%>/de_images/button_help.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doHelp()' title="<%=sTxtHelp%>" class=toolbutton></td>
										  <% end if %>
									  </tr>
									</table>
								  </td>
								</tr>
									<tr>
								  <td class="body" bgcolor="#808080"><img src="<%=DevEditPath%>/de_images/1x1.gif" width="1" height="1"></td>
								</tr>
								<tr>
								  <td class="body" bgcolor="#FFFFFF"><img src="<%=DevEditPath%>/de_images/1x1.gif" width="1" height="1"></td>
								</tr>
								<tr>
								  <td class="body" height=22>
								  <table border="0" cellspacing="0" cellpadding="1" id=toolbar2>
									  <tr id=de>
										<% if e__hideFont <> true then %>
										<td>
										  <select id="fontDrop" onChange="doFont(this[this.selectedIndex].value)" class="Text120" style="border:3px solid #FFFFFF" unselectable="on">
											<%= BuildFontList() %>
										  </select></td>
										<% end if %>
										<% if e__hideSize <> true then %>
										<td>
										  <select id="sizeDrop" onChange="doSize(this[this.selectedIndex].value)" class=Text50 unselectable="on">
											<%= BuildSizeList() %>
										  </select></td>
										<% end if %>
										<% if e__hideFormat <> true then %>
										<td>
										  <select id="formatDrop" onChange="doFormat(this[this.selectedIndex].value)" class="Text70" unselectable="on">
											<option selected><%=sTxtFormat%>
											<option value="<P>">Normal
											<option value="<H1>">Heading 1
											<option value="<H2>">Heading 2
											<option value="<H3>">Heading 3
											<option value="<H4>">Heading 4
											<option value="<H5>">Heading 5
											<option value="<H6>">Heading 6
										  </select></td>
										<% end if %>
										<% if e__hideStyle <> true then %>
										<td>
										  <select id="sStyles" onChange="applyStyle(this[this.selectedIndex].value);foo.focus();this.selectedIndex=0;foo.focus();" class="Text90" unselectable="on" onmouseenter="doStyles()">
											<option selected><%=sTxtStyle%></option>
											<option value=""><%=sTxtNone%></option>
										  </select></td>
										<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										<% end if %>
										<% if e__hideForeColor <> true then

											' Is the user on an XP machine?
											Dim xp
											if (instr(request.serverVariables("HTTP_USER_AGENT"), "Windows NT 5.1")) > 0 Then
												xp = "_xp"
											elseif (instr(request.serverVariables("HTTP_USER_AGENT"), "Windows 98")) > 0 Then
												xp = "_98"
											End if 

										%>

										<td><span id=fontColor><img id=toolbarFont border="0" src="<%=DevEditPath%>/de_images/button_font_color<%=xp%>.gif" width="21" height="20" onmouseover="button_over(this);  button_over(toolbarFontdrop)" onmouseout="button_out(this); button_out(toolbarFontdrop)" onmousedown="button_down(this);" onClick="(isAllowed()) ? doColorDirectly(1) : foo.focus()" class=toolbutton title="<%=sTxtColour%>"></span><img id=toolbarFontdrop class=toolbutton onMouseDown="button_down(this);" onMouseOver="button_over(this); button_over(toolbarFont)" onClick="(isAllowed()) ? showMenu('colorMenu',157,158) : foo.focus()" onMouseOut="button_out(this); button_out(toolbarFont);" width="7" height="20" src="<%=DevEditPath%>/de_images/button_drop_menu.gif" border=0 unselectable="on"></td>
										<% end if %>
										<% if e__hideBackColor <> true then %>
										<td><span id=fontHighlight><img id=toolbarHighlight border="0" src="<%=DevEditPath%>/de_images/button_highlight<%=xp%>.gif" width="21" height="20" onmouseover="button_over(this);  button_over(toolbarHighlightdrop)" onmouseout="button_out(this); button_out(toolbarHighlightdrop)" onmousedown="button_down(this);" onClick="(isAllowed()) ? doColorDirectly(2) : foo.focus()" class=toolbutton title="<%=sTxtBackColour%>"></span><img id=toolbarHighlightdrop class=toolbutton onMouseDown="button_down(this);" onMouseOver="button_over(this); button_over(toolbarHighlight)" onClick="(isAllowed()) ? showMenu('colorMenu2',157,158) : foo.focus()" onMouseOut="button_out(this); button_out(toolbarHighlight);" width="7" height="20" src="<%=DevEditPath%>/de_images/button_drop_menu.gif" border=0 unselectable="on"></td>
										<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										<% end if %>
										<% if e__hideTable <> true then %>
										<td id=toolbarTables>
										  <img border="0" src="<%=DevEditPath%>/de_images/button_table_down.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="(isAllowed()) ? showMenu('tableMenu',160,262) : foo.focus()" class=toolbutton title="<%=sTxtTableFunctions%>"></td>
										<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										<% end if %>
										<% if e__hideForm <> true then %>
										<td>
										  <img class=toolbutton onMouseDown=button_down(this); onMouseOver=button_over(this); onClick="(isAllowed()) ? showMenu('formMenu',180,210) : foo.focus()" onMouseOut=button_out(this); type=image width="21" height="20" src="<%=DevEditPath%>/de_images/button_form_down.gif" border=0 title="<%=sTxtFormFunctions%>"></td>
										<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										<% end if %>
										
										<% if e__hideFlash <> true then %>
										<td>
										  <img border="0" src="<%=DevEditPath%>/de_images/button_flash.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="doFlash()" class=toolbutton title="<%=sTxtFlash%>"></td>
										<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										<% end if %>

										<% if e__hideImage <> true then %>
										<td>
										  <img border="0" src="<%=DevEditPath%>/de_images/button_image.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="doImage()" class=toolbutton title="<%=sTxtImage%>"></td>
										<td><img src="<%=DevEditPath%>/de_images/seperator.gif" width="2" height="21"></td>
										<% end if %>
										<% if e__hideTextBox <> true then %>
										<td>
										  <img border="0" src="<%=DevEditPath%>/de_images/button_textbox.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="doTextbox()" class=toolbutton title="<%=sTxtTextbox%>"></td>
										<% end if %>
										 <% if e__hideHorizontalRule <> true then %>
											<td>
												<img border="0" src="<%=DevEditPath%>/de_images/button_hr.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick='doCommand("InsertHorizontalRule");foo.focus();' title="<%=sTxtInsertHR%>" class=toolbutton></td>
										  <% end if %>
										<% if e__hideSymbols <> true then %>
										<td>
										  <img border="0" src="<%=DevEditPath%>/de_images/button_chars.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="doChars()" class=toolbutton title="<%=sTxtChars%>"></td>
										<% end if %>
										<% if e__hideProps <> true then %>
										<td>
										  <img border="0" src="<%=DevEditPath%>/de_images/button_properties.gif" width="21" height="20" onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" onClick="ModifyProperties()" class=toolbutton title="<%=sTxtPageProperties%>"></td>
										<% end if %>

										<% if e__hideClean <> true then %>
										<td>				
										  <img class=toolbutton onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="cleanCode()" onmouseout="button_out(this);" type=image width="21" height="20" src="<%=DevEditPath%>/de_images/button_clean_code.gif" border=0 title="<%=sTxtCleanCode%>"></td>
										<% end if %>

										<% if e__hasCustomInserts = true then %>
										<td>
											<img class=toolbutton onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="doCustomInserts()" onmouseout="button_out(this);" type=image width="21" height="20" src="<%=DevEditPath%>/de_images/button_custom_inserts.gif" border=0 title="<%=sTxtCustomInserts%>"></td>
										<% end if %>

										<% if e__hideAbsolute <> true then %>
										<td>
											<img id="fontAbsolutePosition_off" disabled class=toolbutton onmousedown="button_down(this);" onmouseover="button_over(this);" width="21" height="20" src="<%=DevEditPath%>/de_images/button_absolute_disabled.gif" border=0 title="<%=sTxtTogglePosition%>"><img id="fontAbsolutePosition" class=toolbutton onmousedown="button_down(this);" onmouseover="button_over(this);" onClick="doCommand('AbsolutePosition')" onmouseout="button_out2(this);" type=image width="21" height="20" src="<%=DevEditPath%>/de_images/button_absolute.gif" border=0 title="<%=sTxtTogglePosition%>" style="display:none"></td>
										<% end if %>

										<% if e__hideGuidelines <> true then %>
										<td>
										  <img class=toolbutton onMouseDown="button_down(this);" onMouseOver="button_over(this);" onClick="toggleBorders()" onMouseOut="button_out2(this);" type=image width="21" height="20" src="<%=DevEditPath%>/de_images/button_show_borders.gif" border=0 title="<%=sTxtToggleGuidelines%>" id=guidelines></td>
										<% end if %>
										<td><div class="pasteArea" id="myTempArea" contentEditable></div></td>
									  </tr>
									</table>
								  </td>
								</tr>
							  </table>
							</td>
						  </tr> 
						</table>
						<!-- table menu -->
						<DIV ID="tableMenu" STYLE="display:none">
						<table border="0" cellspacing="0" cellpadding="0" width=160 style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
						  <tr onClick="parent.ShowInsertTable()" title="<%=sTxtTable%>" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=20> 
							  &nbsp;&nbsp;<%=sTxtTable%>...&nbsp; </td>
						  </tr>
						  <tr onClick=parent.ModifyTable(); title="<%=sTxtTableModify%>" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=modifyTable> 
							  &nbsp;&nbsp;<%=sTxtTableModify%>...&nbsp;</td>
						  </tr>
						  <tr title="<%=sTxtCellModify%>" onClick=parent.ModifyCell() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=modifyCell> 
							&nbsp;&nbsp;<%=sTxtCellModify%>...&nbsp; </td>
						  </tr>
						  <tr height=10> 
							<td align=center><img src="<%=DevEditPath%>/de_images/vertical_spacer.gif" width="140" height="2"></td>
						  </tr>
						  <tr title="<%=sTxtInsertColA%>" onClick=parent.InsertColAfter() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=colAfter> 
							  &nbsp;&nbsp;<%=sTxtInsertColA%>&nbsp;
							</td>
						  </tr>
						  <tr title="<%=sTxtInsertColB%>" onClick=parent.InsertColBefore() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=colBefore> 
							  &nbsp;&nbsp;<%=sTxtInsertColB%>&nbsp;
							</td>
						  </tr>
						  <tr height=10> 
							<td align=center><img src="<%=DevEditPath%>/de_images/vertical_spacer.gif" width="140" height="2"></td>
						  </tr>
						  <tr title="<%=sTxtInsertRowA%>" onClick=parent.InsertRowAbove() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=rowAbove> 
							  &nbsp;&nbsp;<%=sTxtInsertRowA%>&nbsp;
							</td>
						  </tr>
						  <tr title="<%=sTxtInsertRowB%>" onClick=parent.InsertRowBelow() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=rowBelow> 
							  &nbsp;&nbsp;<%=sTxtInsertRowB%>&nbsp;
							</td>
						  </tr>
						  <tr height=10> 
							<td align=center><img src="<%=DevEditPath%>/de_images/vertical_spacer.gif" width="140" height="2"></td>
						  </tr>
						  <tr title="<%=sTxtDeleteRow%>" onClick=parent.DeleteRow() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=deleteRow>
							  &nbsp;&nbsp;<%=sTxtDeleteRow%>&nbsp;
							</td>
						  </tr>
						  <tr title="<%=sTxtDeleteCol%>" onClick=parent.DeleteCol() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=deleteCol>
							  &nbsp;&nbsp;<%=sTxtDeleteCol%>&nbsp;
							</td>
						  </tr>
						  <tr height=10> 
							<td align=center><img src="<%=DevEditPath%>/de_images/vertical_spacer.gif" width="140" height="2" tabindex=1 HIDEFOCUS></td>
						  </tr>
						  <tr title="<%=sTxtIncreaseColSpan%>" onClick=parent.IncreaseColspan() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=increaseSpan>
							  &nbsp;&nbsp;<%=sTxtIncreaseColSpan%>&nbsp;
							</td>
						  </tr>
						  <tr title="<%=sTxtDecreaseColSpan%>" onClick=parent.DecreaseColspan() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=20 id=decreaseSpan>
							  &nbsp;&nbsp;<%=sTxtDecreaseColSpan%>&nbsp;
							</td>
						  </tr>
						</table>
						</div>
						<!-- end table menu -->

						<!-- form menu -->
						<DIV ID="formMenu" STYLE="display:none;">
						<table border="0" cellspacing="0" cellpadding="0" width=180 style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
						  <tr title="<%=sTxtForm%>" onClick=parent.insertForm() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22>
							  <img width="21" height="20" src="<%=DevEditPath%>/de_images/button_form.gif" border=0 align="absmiddle">&nbsp;<%=sTxtForm%>...&nbsp;</td>
						  </tr>
						  <tr title="<%=sTxtFormModify%>" onClick=parent.modifyForm() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" id="modifyForm1" height=22 class=dropDown>
							  <img id="modifyForm2" width="21" height="20" src="<%=DevEditPath%>/de_images/button_modify_form.gif" border=0 align="absmiddle">&nbsp;<%=sTxtFormModify%>...&nbsp;</td>
						  </tr>
						  <tr height=10> 
							<td align=center><img src="<%=DevEditPath%>/de_images/vertical_spacer.gif" width="140" height="2" tabindex=1 HIDEFOCUS></td>
						  </tr>
						  <tr title="<%=sTxtTextField%>" onClick=parent.doTextField() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22>
							  <img width="21" height="20" src="<%=DevEditPath%>/de_images/button_textfield.gif" border=0 align="absmiddle">&nbsp;<%=sTxtTextField%>...&nbsp;</td>
						  </tr>
						  <tr title="<%=sTxtTextArea%>" onClick=parent.doTextArea() onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=22>
							  <img type=image width="21" height="20" src="<%=DevEditPath%>/de_images/button_textarea.gif" border=0 align="absmiddle">&nbsp;<%=sTxtTextArea%>...&nbsp;</td>
						  </tr>
						  <tr title="<%=sTxtHidden%>" onClick=parent.doHidden(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							<td style="cursor: hand; font:8pt tahoma;" height=22>
							  <img width="21" height="20" src="<%=DevEditPath%>/de_images/button_hidden.gif" border=0 align="absmiddle">&nbsp;<%=sTxtHidden%>...&nbsp;</td>
						  </tr>
						  <tr title="<%=sTxtButton%>" onClick=parent.doButton(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22>
							  <img width="21" height="20" src="<%=DevEditPath%>/de_images/button_button.gif" border=0 align="absmiddle">&nbsp;<%=sTxtButton%>...&nbsp;</td>
						  </tr>
						  <tr title="<%=sTxtCheckbox%>" onClick=parent.doCheckbox(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22>
							  <img width="21" height="20" src="<%=DevEditPath%>/de_images/button_checkbox.gif" border=0 align="absmiddle">&nbsp;<%=sTxtCheckbox%>...&nbsp;</td>
						  </tr>
						  <tr title="<%=sTxtRadioButton%>" onClick=parent.doRadio(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22>
							  <img width="21" height="20" src="<%=DevEditPath%>/de_images/button_radio.gif" border=0 align="absmiddle">&nbsp;<%=sTxtRadioButton%>...&nbsp;</td>
						  </tr>
						  <tr title="<%=sTxtSelect%>" onClick=parent.doSelect(); onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22>
							  <img width="21" height="20" src="<%=DevEditPath%>/de_images/button_select.gif" border=0 align="absmiddle">&nbsp;<%=sTxtSelect%>...&nbsp;</td>
						  </tr>
						</table>
						</div>
						<!-- formMenu -->

						<!-- zoom menu -->
						<DIV ID="zoomMenu" STYLE="display:none;">
						<table border="0" cellspacing="0" cellpadding="0" width=65 style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
						  <tr onClick=parent.doZoom(500) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom500_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom500_,0);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom500_">
							 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;500%&nbsp;</td>
						  </tr>
						  <tr onClick=parent.doZoom(200) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom200_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom200_,0);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom200_">
							  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;200%&nbsp;</td>
						  </tr>
						  <tr onClick=parent.doZoom(150) onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom150_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom150_,0);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom150_">
							  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;150%&nbsp;</td>
						  </tr>
						  <tr onClick="parent.doZoom(100)" onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom100_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom100_,0)";">
							<td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom100_">
							  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;100%&nbsp;</td>
						  </tr>
						  <tr onClick=parent.doZoom(75); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom75_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom75_,0);">
							<td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom75_">
							  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;75%&nbsp;</td>
						  </tr>
						  <tr onClick=parent.doZoom(50); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom50_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom50_,0);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom50_">
							  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;50%&nbsp;</td>
						  </tr>
						  <tr onClick=parent.doZoom(25); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom25_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom25_,0);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom25_">
							  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;25%&nbsp;</td>
						  </tr>
						  <tr onClick=parent.doZoom(10); onMouseOver="parent.contextHilite(this); parent.toggleTick(zoom10_,1);" onMouseOut="parent.contextDelite(this); parent.toggleTick(zoom10_,0);"> 
							<td style="cursor: hand; font:8pt tahoma;" height=22 id="zoom10_">
							  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10%&nbsp;</td>
						  </tr>
						</table>
						</div>
						<!-- zoomMenu -->

						<DIV ID="colorMenu" STYLE="display:none;">
						<table cellpadding="0" cellspacing="5" style="cursor: hand;font-family: Verdana; font-size: 6px; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface"><tr><td>

						<table cellpadding=0 cellspacing=0 style="font-size: 3px;">
						  <tr>
							<td colspan="10" style="border: solid 1px #d4d0c8;" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" onClick="parent.doColor()"><div style="border: solid 1px #808080; padding: 2px; margin: 2px;">
							<table cellspacing=0 cellpadding=0 border=0 width=90% style="font-size:3px">
							<tr><td><div style="background-color:#000000; border:solid 1px #808080; width:12px; height:12px"></div></td><td align=center style="font-size:11px"><%=sTxtNone%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
							</table>
							</div>
							</td>
						  </tr>
						  <tr><td>&nbsp;</td></tr>
						  <tr>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#000000; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#993300; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#333300; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#003300; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#003366; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#000099; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#333399; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#333333; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
						  </tr>
						  <tr>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#990000; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FF6600; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#999900; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#009900; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#009999; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#0000FF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#666699; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#808080; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
						  </tr>
						  <tr>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FF0000; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FF9900; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#99CC00; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#339966; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#33CCCC; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#3366FF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#990099; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#999999; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
						  </tr><tr>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FF00FF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FFCC00; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FFFF00; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#00FF00; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#00FFFF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#00CCFF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#993366; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#CCCCCC; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
						  </tr>
						  <tr>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FF99CC; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FFCC99; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FFFF99; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#CCFFCC; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#CCFFFF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#99CCFF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#CC99FF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
							<td onClick="parent.doColor(this)" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" style="padding:2px;border:solid 1px #d4d0c8;"><div style="background-color:#FFFFFF; border:solid 1px #808080; width:12px; height:12px">&nbsp;</div></td>
						  </tr>
						  <tr><td>&nbsp;</td></tr>
						  <tr>
							<td colspan="10" style="height:23px; font-family: arial; font-size:11px; border: solid 1px #d4d0c8;" onMouseOver="parent.button_over(this)" onMouseOut="parent.button_out(this)" onClick="parent.doMoreColors()" align=center>&nbsp;<%=sTxtMoreColors%>...</td>
						  </tr>
						</table>

						</td></tr>
						</table>
						</DIV>


						<DIV ID="contextMenu" style="display:none;">
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
						  <tr id=cmCut onClick ='parent.document.execCommand("Cut");parent.oPopup2.hide()'>
							<td style="cursor:default; font:8pt tahoma; border: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCut%>&nbsp;</td>
						  </tr>
						  <tr id=cmCopy onClick ='parent.document.execCommand("Copy");parent.oPopup2.hide()'> 
							<td style="cursor:default; font:8pt tahoma; border: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCopy%>&nbsp;</td>
						  </tr>
						  <tr id=cmPaste onClick ='parent.document.execCommand("Paste");parent.oPopup2.hide()'> 
							<td style="cursor:default; font:8pt tahoma; border: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPaste%>&nbsp;</td>
						  </tr>
						</table>
						</div>

						<DIV ID="cmTableMenu" style="display:none">
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.ModifyTable();'> 
							<td style="cursor:default; font:8pt tahoma; border: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtTableModify%>...&nbsp;</td>
						  </tr>
						</table>
						</DIV>

						<DIV ID="cmTableFunctions" style="display:none">
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.ModifyCell();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtCellModify%>...&nbsp;</td>
						  </tr>
						</table>
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.InsertColBefore(); parent.oPopup2.hide();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertColB%>&nbsp;</td>
						  </tr>
						  <tr onClick ='parent.InsertColAfter(); parent.oPopup2.hide();'> 
						   <td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertColA%>&nbsp;</td>
						  </tr>
						</table>
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.InsertRowAbove(); parent.oPopup2.hide();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertRowA%>&nbsp;</td>
						  </tr>
						  <tr onClick ='parent.InsertRowBelow(); parent.oPopup2.hide();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtInsertRowB%>&nbsp;</td>
						  </tr>
						</table>
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.DeleteRow(); parent.oPopup2.hide();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtDeleteRow%>&nbsp;</td>
						  </tr>
						  <tr onClick ='parent.DeleteCol(); parent.oPopup2.hide();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtDeleteCol%>&nbsp;</td>
						  </tr>
						</table>
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.IncreaseColspan(); parent.oPopup2.hide();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtIncreaseColSpan%>&nbsp;</td>
						  </tr>
						  <tr onClick ='parent.DecreaseColspan(); parent.oPopup2.hide();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp<%=sTxtDecreaseColSpan%>&nbsp;</td>
						  </tr>
						</table>
						</DIV>

						<DIV ID="cmImageMenu" style="display:none">
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.doImage();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtModifyImage%>...&nbsp;</td>
						  </tr>
						</table>
						</DIV>

						<DIV ID="cmLinkMenu" style="display:none">
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.doLink();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtHyperLink%>...&nbsp;</td>
						  </tr>
						</table>
						</DIV>

						<DIV ID="cmSpellMenu" style="display:none">
						<table border="0" cellspacing="0" cellpadding="3" width="<%=sTxtContextMenuWidth-2%>" style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: #808080 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: #808080 1px solid;" bgcolor="threedface">
						  <tr onClick ='parent.spellCheck();'> 
							<td style="cursor:default; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" class="parent.toolbutton" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);">
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;Check Spelling...&nbsp;</td>
						  </tr>
						</table>
						</DIV>

						<!-- Start Paste Menu -->
						<DIV ID="pasteMenu" STYLE="display:none">
						<table border="0" cellspacing="0" cellpadding="0" width=180 style="BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;" bgcolor="threedface">
						  <tr onClick="parent.doCommand('Paste');"> 
							<td height=20 style="cursor: hand; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
						&nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPaste%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ctrl+V </td>
						  </tr>
						  <tr onClick="parent.pasteWord();"> 
							<td height=20 style="cursor: hand; font:8pt tahoma; BORDER-LEFT: threedface 1px solid; BORDER-RIGHT: threedface 1px solid; BORDER-TOP: threedface 1px solid; BORDER-BOTTOM: threedface 1px solid;" onMouseOver="parent.contextHilite(this);" onMouseOut="parent.contextDelite(this);"> 
							  &nbsp&nbsp;&nbsp;&nbsp&nbsp;<%=sTxtPasteWord%>&nbsp;&nbsp;&nbsp;&nbsp;Ctrl+D </td>
						  </tr>
						</table>
						</div>
						<!-- End Paste Menu -->
					<!-- END DEVEDIT TOOLBAR -->
        							
					<!-- removed security=restricted to allow Flash insertion -->
					<span id=divContainer style="width:100%" class=divContainer>
					<iFrame onBlur="updateValue()" contenteditable id="foo" src='' style="width:100%; border: 1px solid #D4D0C8"></iFrame>
					<iFrame id="fooStyles" src='' style="display:none"></iFrame>
					<iframe onBlur="updateValue()" id=previewFrame height=80% style="width=100%; display:none" style="border: 1px solid #D4D0C8"></iframe>
					</span>
        						<script>
								document.getElementById("foo").style.setExpression("height", "document.body.clientHeight - 85")
								document.getElementById("previewFrame").style.setExpression("height", "document.body.clientHeight - 59")
								</script>

        							<table cellpadding=0 cellspacing=0 width=100% style="background-color: threedface">
        								<tr>
        									<td height=22><img style="cursor:hand;" id=editTab src=<%=DevEditPath%>/de_images/status_edit_up.gif width=98 height=22 border=0 onClick=editMe()><img style="cursor:hand; <% if e__disableSourceMode = true then %>display:none<% end if %>" id=sourceTab src=<%=DevEditPath%>/de_images/status_source.gif width=98 height=22 border=0 onClick=sourceMe()><img style="cursor:hand; <% if e__disablePreviewMode = true then %>display:none<% end if %>" id=previewTab src=<%=DevEditPath%>/de_images/status_preview.gif width=98 height=22 border=0 onClick=previewMe()></td><td width=100%  height=22 background=<%=DevEditPath%>/de_images/status_border.gif>&nbsp;</td><td background=<%=DevEditPath%>/de_images/status_border.gif id=statusbar align=right valign=bottom><img src=<%=DevEditPath%>/de_images/button_zoom.gif width=42 height=17 valign=bottom onmouseover="button_over(this);" onmouseout="button_out(this);" onmousedown="button_down(this);" class=toolbutton onClick="showMenu('zoomMenu',65,178)"></td>
        								</tr>
        							</table>
						</span>
						<!-- end mod -->
        
        				<script language="JavaScript">
        							
        					var fooWidth = "<%=e__controlWidth %>";
        					var fooHeight = "<%=e__controlHeight %>";
        
        					function setValue()
							{
									foo.document.write('<%=Replace(Replace(e__initialValue, "</script>", "<\/script>", 1, -1, 1), "<", "&lt;") %>');
									foo.document.close();

									fooStyles.document.write('<%=Replace(Replace(e__initialValueNoBase, "</script>", "<\/script>", 1, -1, 1), "<", "&lt;") %>');
									fooStyles.document.close();
        					}
        
							function updateValue(isSave)
							{ 
							 if (document.activeElement) {
							  if ((document.activeElement.parentElement.id == "de") && (!isSave)) {
							   return false;
							  } else {
							   if (parent.document.all.<%=e__controlName%>_html != null) {
								parent.document.all.<%=e__controlName%>_html.value = SaveHTMLPage();
							   }
							  }
							 }
							}

        				</script>
        
        			<%

					'End the iFrame source text area buffer
					%>
						</textarea>

						<iframe id="<%=e__controlName%>_frame" width="<%=e__controlWidth%>" height="<%=e__controlHeight%>" frameborder=0 scrolling="no"></iframe>

						<input type="hidden" name="<%=e__controlName%>_html">

						<script language="JavaScript">
							<% if DevEditError = true then %>
								<%=e__controlName%>_frame.document.write("<b>Error:</b> The value you have specified for the setDevEditPath function is invalid.")
								<%=e__controlName%>_frame.document.close()
								</script>
							<%	response.end
								End if
							%>
							
							<%=e__controlName%>_frame.document.write(document.getElementById("<%=e__controlName%>_src").value)
							<%=e__controlName%>_frame.document.close()
							<%=e__controlName%>_frame.document.body.style.margin = "0px";

							// set the width of the editor.
							// dont forget to add the toolbar2 id in the toolbar
							
							setTimeout("doMyWidth();",100);

							function doMyWidth() {
								minWidth = 0
								if (<%=e__controlName%>_frame.document.getElementById("toolbar1").clientWidth > <%=e__controlName%>_frame.document.getElementById("toolbar2").clientWidth)
									minWidth = <%=e__controlName%>_frame.document.getElementById("toolbar1").clientWidth
								else
									minWidth = <%=e__controlName%>_frame.document.getElementById("toolbar2").clientWidth

								document.getElementById("<%=e__controlName%>_frame").style.setExpression("width", "setFrameWidth_<%=e__controlName%>()")
							}

							function setFrameWidth_<%=e__controlName%> () {

							frame_offsetLeft = document.getElementById("<%=e__controlName%>_frame").offsetLeft

								if (document.body.clientWidth > minWidth) {
									if ("<%=e__controlWidth%>".indexOf('%') > -1) {

										if (parseInt("<%=e__controlWidth%>") / 100 * document.body.clientWidth < (minWidth + frame_offsetLeft + 11))
											return minWidth + 2
										else
											return "<%=e__controlWidth%>"

									} else {

										if (parseInt("<%=e__controlWidth%>") < minWidth + 2)
											return minWidth + 2
										else
											return "<%=e__controlWidth%>"
									}

								} else {
									return minWidth + 2
								}
							}

						</script>

					<%
			end if
			
		end sub
		
	end class

	private function PageInsertImage()

		On Error Resume next

		dt = Request("dt")
		tn = Request("tn")
		dd = Request("dd")
		du = Request("du")
		deveditPath = Request("DEP")

		'Get a list of images from the directory
		set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		validImageTypes = array("image/pjpeg", "image/jpeg", "image/gif", "image/x-png", "image/tiff", "image/tif", "image/bmp", "image/x-MS-bmp")

		' Added for v5.0
		validImageExts = array("gif","jpg","jpeg","bmp","tiff","tif","png")
		isValidExt = false

		ImageDirectory = Request("imgDir")

		HideWebImage = Request("wi")
		Response.Write HiddenWebImage
		URL = Request.ServerVariables("http_host")
		scriptName = DevEditPath & "/class.devedit.asp"

		'Do we need to kill the following slash?
		serverName = URL
		'if right(serverName, 1) = "/" then
			'serverName = left(serverName, len(serverName)-1)
		'end if

		'Workout the location of class.devedit.asp
		scriptDir = strreverse(Request.ServerVariables("path_info"))
		slashPos = instr(1, scriptDir, "/")
		scriptDir = strreverse(mid(scriptDir, slashPos, len(scriptDir)))
		
		if Request("imgSrc") <> "" then
		
			'Delete the image
			imgPath = ImageDirectory & "/" & Request("imgSrc")
			imgExt = strreverse(left(strreverse(Request("imgSrc")), instr(strreverse(Request("imgSrc")), ".") -1))
			isValidExt = false

			'Does the image have a valid extension?
			for i = 0 to uBound(validImageExts)
				if ucase(imgExt) = ucase(validImageExts(i)) then
					isValidExt = true
				end if
			next

			if isValidExt = true then
				deleteSuccess = objFSO.DeleteFile(Server.MapPath(imgPath))

				if Err.Number <> 0 then
					statusText = sTxtCantDelete
				else
					statusText = sTxtImageDeleted
				end if
			else
				statusText = sTxtInvalidImageType
			end if

		end if
		
		if request.querystring("ToDo") = "UploadImage" then

			'Get DT from the querystring
			if Instr(Request("QUERY_STRING"), "dt=0") > 0 then
				dt = 0
			else
				dt = 1
			end if

			if Instr(Request("QUERY_STRING"), "wi=0") > 0 then
				HideWebImage = 0
			else
				HideWebImage = 1
			end if

			if Instr(Request("QUERY_STRING"), "dd=0") > 0 then
				dd = 0
			else
				dd = 1
			end if

			if Instr(Request("QUERY_STRING"), "du=0") > 0 then
				du = 0
			else
				du = 1
			end if

			if Instr(Request("QUERY_STRING"), "wi=0") > 0 then
				wi = 0
			else
				wi = 1
			end if

			if Instr(Request("QUERY_STRING"), "tn=0") > 0 then
				tn = 0
			else
				tn = 1
			end if

			'Upload the image to the images directory
			set objFile = new Loader
			objFile.Initialize

			validFileType = false
			errorText = ""

			'Data for first file upload
			newFileName = objFile.GetFileName("upload")
			newFileType = objFile.GetContentType("upload")
			newFileData = objFile.GetFileData("upload")

			'Data for second file upload
			newFileName1 = objFile.GetFileName("upload1")
			newFileType1 = objFile.GetContentType("upload1")
			newFileData1 = objFile.GetFileData("upload1")

			'Data for third file upload
			newFileName2 = objFile.GetFileName("upload2")
			newFileType2 = objFile.GetContentType("upload2")
			newFileData2 = objFile.GetFileData("upload2")

			'Data for fourth file upload
			newFileName3 = objFile.GetFileName("upload3")
			newFileType3 = objFile.GetContentType("upload3")
			newFileData3 = objFile.GetFileData("upload3")

			'Data for fifth file upload
			newFileName4 = objFile.GetFileName("upload4")
			newFileType4 = objFile.GetContentType("upload4")
			newFileData4 = objFile.GetFileData("upload4")

			numImages = 0

			'---------------------------------------------------------
			'Is the first image a valid file type?
			validFileType = false

			if newFileName = "" then
				isBlank = 1
			else
				isBlank = 0
			end if

			if isBlank = 0 then
				for i = 0 to uBound(validImageTypes)
					if(newFileType = validImageTypes(i)) then
						validFileType = true
					end if
				next

				if validFileType = false then
					'Invalid file type
					statusText = sTxtImageErr
				else
					'Does this image already exist? We don't want to overwrite it
					uploadSuccess = objFile.SaveToFile("upload", Server.MapPath(ImageDirectory & "/" & newFileName))

					if Err.Number <> 0 then
						statusText = sTxtCantUpload
					else
						statusText = newFileName & " " & sTxtUploadSuccess & "!"	
						numImages = numImages + 1
					end if

				end if
			end if

			'---------------------------------------------------------
			'Is the second image a valid file type?
			validFileType = false

			if newFileName1 = "" then
				isBlank = 1
			else
				isBlank = 0
			end if

			if isBlank = 0 then
				for i = 0 to uBound(validImageTypes)
					if(newFileType1 = validImageTypes(i)) then
						validFileType = true
					end if
				next

				if validFileType = false then
					'Invalid file type
					statusText = sTxtImageErr
				else
					uploadSuccess = objFile.SaveToFile("upload1", Server.MapPath(ImageDirectory & "/" & newFileName1))

					if Err.Number <> 0 then
						statusText = sTxtCantUpload
					else
						statusText = newFileName & " " & sTxtUploadSuccess & "!"	
						numImages = numImages + 1
					end if

				end if
			end if

			'---------------------------------------------------------
			'Is the third image a valid file type?
			validFileType = false

			if newFileName2 = "" then
				isBlank = 1
			else
				isBlank = 0
			end if

			if isBlank = 0 then
				for i = 0 to uBound(validImageTypes)
					if(newFileType2 = validImageTypes(i)) then
						validFileType = true
					end if
				next
			
				if validFileType = false then
					'Invalid file type
					statusText = sTxtImageErr
				else
					uploadSuccess = objFile.SaveToFile("upload2", Server.MapPath(ImageDirectory & "/" & newFileName2))

					if Err.Number <> 0 then
						statusText = sTxtCantUpload
					else
						statusText = newFileName & " " & sTxtUploadSuccess & "!"	
						numImages = numImages + 1
					end if

				end if
			end if		


			'---------------------------------------------------------
			'Is the fourth image a valid file type?
			validFileType = false

			if newFileName3 = "" then
				isBlank = 1
			else
				isBlank = 0
			end if

			if isBlank = 0 then
				for i = 0 to uBound(validImageTypes)
					if(newFileType3 = validImageTypes(i)) then
						validFileType = true
					end if
				next
			
				if validFileType = false then
					'Invalid file type
					statusText = sTxtImageErr
				else
					uploadSuccess = objFile.SaveToFile("upload3", Server.MapPath(ImageDirectory & "/" & newFileName3))

					if Err.Number <> 0 then
						statusText = sTxtCantUpload
					else
						statusText = newFileName & " " & sTxtUploadSuccess & "!"	
						numImages = numImages + 1
					end if

				end if
			end if


			'---------------------------------------------------------
			'Is the fifth image a valid file type?
			validFileType = false

			if newFileName4 = "" then
				isBlank = 1
			else
				isBlank = 0
			end if

			if isBlank = 0 then
				for i = 0 to uBound(validImageTypes)
					if(newFileType4 = validImageTypes(i)) then
						validFileType = true
					end if
				next
			
				if validFileType = false then
					'Invalid file type
					statusText = sTxtImageErr
				else
					uploadSuccess = objFile.SaveToFile("upload4", Server.MapPath(ImageDirectory & "/" & newFileName4))

					if Err.Number <> 0 then
						statusText = sTxtCantUpload
					else
						statusText = newFileName & " " & sTxtUploadSuccess & "!"	
						numImages = numImages + 1
					end if

				end if
			end if

		end if

		'Where should we get the DEP variable from?
		DEP = Request("DEP")

		if numImages > 1 then
			statusText = sTxtImageSuccess
		end if

		SImageDirectory = Server.MapPath(ImageDirectory)

		If (objFSO.FolderExists(SImageDirectory) = true) Then
			set objImageDir = objFSO.GetFolder(SImageDirectory)
		else
			response.write sTxtImageDirNotConfigured
			response.end
		End if

		' set objImageDir = objFSO.GetFolder(SImageDirectory)
		set objImageList = objImageDir.Files
		counter = 0

		'Get all images into a JavaScript array so that we can workout whether or not
		'uploading an image would overwrite an existing one
		imageJS = "var imageDir = Array("

		for each imageFile in objImageList
			imageJS = imageJS & "'" & imageFile.name & "',"
		next

		if right(imageJS, 1) = "," then
			imageJS = mid(imageJS, 1, len(imageJS)-1)
		end if

		imageJS = imageJS & ")"
	%>

	<script language=JavaScript>

	window.onload = function()
	{
		document.getElementById("DEP").value = window.opener.deveditPath;
	}
	
	window.opener.doStyles()
	window.onerror = function() { return true; }
	var myPage = window.opener;
	var imageAlign

	<%=imageJS%>

	if (window.opener.imageEdit) {
		imageAlign = window.opener.selectedImage.align
	}

	function toggleUploadDiv()
	{
		if(uploadDiv.style.display == "none")
		{
			document.getElementById("toggleButton").value = "«";
			document.getElementById("upload").focus();
			document.getElementById("upload").select();
			uploadDiv.style.display = "inline";
			dummyDiv.style.display = "inline";
			divList.style.height = 225;
			previewWindow.style.height = 50;
		}
		else
		{
			document.getElementById("toggleButton").value = "»";
			document.getElementById("upload").focus();
			document.getElementById("upload").select();
			uploadDiv.style.display = "none";
			dummyDiv.style.display = "none";
			divList.style.height = 325;
			previewWindow.style.height = 150;
		}
	}

	function outputImageLibraryOptions()
	{
		document.write(opener.imageLibs);

		// Loop through all of the image libraries and find the selected one
		for(i = 0; i < selImageLib.options.length; i++)
		{
			if(selImageLib.options[i].value == "<%=ImageDirectory%>")
			{
				selImageLib.selectedIndex = i;
				break;
			}
		}
	}

	function switchImageLibrary(thePath)
	{
		// Change the path of the image library
		DEP = document.getElementById("DEP").value;
		document.location.href = '<%=HTTPStr%>://<%=serverName%><%=scriptDir%>class.devedit.asp?ToDo=InsertImage&DEP='+DEP+'&imgDir='+thePath+'&dd=<%=dd%>&du=<%=du%>&wi=<%=wi%>&tn=<%=tn%>&dt=<%=dt%>&wi=<%=HideWebImage%>';
	}

	function previewModify() {

		var imageWidth = myPage.selectedImage.width;
		var imageHeight = myPage.selectedImage.height;
		var imageBorder = myPage.selectedImage.border;
		var imageAltTag = myPage.selectedImage.alt;
		var imageHspace = myPage.selectedImage.hspace;
		var imageVspace = myPage.selectedImage.vspace;

		document.getElementById("previewWindow").innerHTML = "<img src=" + selectedImage.replace(/ /g, "%20") + ">"

		insertButton.value = "<%=sTxtImageModify%>"
		document.title = "<%=sTxtModifyImage%>"
		if (document.getElementById("deleteButton") != null) {
			deleteButton.disabled = true
		}

		previewButton.disabled = false
		insertButton.disabled = false

		if (document.getElementById("backgdButton") != null) {
			backgdButton.disabled = false
		}

		image_width.value = imageWidth;
		image_height.value = imageHeight;

		if (imageBorder == "") {
			imageBorder = "0"
		}

		border.value = imageBorder;
		alt_tag.value = imageAltTag;
		hspace.value = imageHspace;
		vspace.value = imageVspace;
		// tableForm.cell_width.value = cellWidth;
		this.focus();
	}

	function deleteImage(imgSrc)
	{
		var delImg = confirm("<%=sTxtImageDelete%>");
		DEP = document.getElementById("DEP").value;

		if (delImg == true) {
			document.location.href = '<%=HTTPStr%>://<%=serverName%><%=scriptDir%>class.devedit.asp?ToDo=DeleteImage&DEP='+DEP+'&imgDir=<%=ImageDirectory%>&tn=<%=tn%>&dt=<%=dt%>&wi=<%=HideWebImage%>&du=<%=du%>&imgSrc='+imgSrc;
		}

	}

	function setBackground(imgSrc)
	{
		var setBg = confirm("<%=sTxtImageSetBackgd%>?");

		if (setBg == true) {
			window.opener.setBackgd(imgSrc);
			self.close();
		}
	}

	function viewImage(imgSrc)
	{
		var sWidth =  screen.availWidth;
		var sHeight = screen.availHeight;
		
		window.open(imgSrc, 'image', 'width=500, height=500,scrollbars=yes,resizable=yes,left='+(sWidth/2-250)+',top='+(sHeight/2-250));
	}

	function grey(tr) {
			tr.className = 'b4';
	}

	function ungrey(tr) {
			tr.className = '';
	}

	function insertImage(imgSrc) {

		var error = 0;

			imageWidth = image_width.value
			imageHeight = image_height.value
			imageBorder = border.value
			imageHspace = hspace.value
			imageVspace = vspace.value

			if (isNaN(imageWidth) || imageWidth < 0) {
				alert("<%=sTxtImageWidthErr%>")
				error = 1
				image_width.select()
				image_width.focus()
			} else if (isNaN(imageHeight) || imageHeight < 0) {
				alert("<%=sTxtImageHeightErr%>")
				error = 1
				image_height.select()
				image_height.focus()
			} else if (isNaN(imageBorder) || imageBorder < 0 || imageBorder == "") {
				alert("<%=sTxtImageBorderErr%>")
				error = 1
				border.select()
				border.focus()
			} else if (isNaN(imageHspace) || imageHspace < 0) {
				alert("<%=sTxtHorizontalSpacingErr%>")
				error = 1
				hspace.select()
				hspace.focus()
			} else if (isNaN(vspace.value) || vspace.value < 0) {
				alert("<%=sTxtVerticalSpacingErr%>")
				error = 1
				vspace.select()
				vspace.focus()
			}

			if (error != 1) {

				var sel = window.opener.foo.document.selection;
				if (sel!=null) {
					var rng = sel.createRange();
					if (rng!=null) {

						if (window.opener.imageEdit) {
							oImage = window.opener.selectedImage
							oImage.src = imgSrc
						} else { 
							HTMLTextField = '<img id="de_element_image" src="' + imgSrc + '">';
							rng.pasteHTML(HTMLTextField)

							oImage = window.opener.foo.document.getElementById("de_element_image")
						}

						if (imageWidth != "")
							oImage.width = imageWidth

						if (imageHeight != "")
							oImage.height = imageHeight

						oImage.alt = alt_tag.value
						oImage.border = border.value
						
						if (hspace.value != "") {
							oImage.hspace = hspace.value
						}

						if (vspace.value != "") {
							oImage.vspace = vspace.value
						} else {
							oImage.removeAttribute('vspace',0)
						}

						if (align[align.selectedIndex].text != "None") {
							oImage.align = align[align.selectedIndex].text
						} else {
							oImage.removeAttribute('align',0)
						}

						styles = sStyles[sStyles.selectedIndex].text

						if (styles != "") {
							oImage.className = styles
						} else {
							oImage.removeAttribute('className',0)
						}

						// window.opener.doToolbar()
						// window.opener.foo.focus();
						self.close();

						if (window.opener.imageEdit) {
							// do nothing
						} else { 
							oImage.removeAttribute("id")
						}


					} // End if
				} // End If
			}
	} // End function

	function insertExtImage() {
		selectedImage = document.getElementById("externalImage").value
		
		if (previousImage != null) {
			previousImage.style.border = "3px solid #FFFFFF"
		}

		document.getElementById("previewWindow").innerHTML = "<img src=" + selectedImage.replace(/ /g, "%20") + ">"

		if (document.getElementById("deleteButton") != null) {
			deleteButton.disabled = true
		}

		previewButton.disabled = false
		insertButton.disabled = false

		if (document.getElementById("backgdButton") != null) {
			backgdButton.disabled = false
		}

	} // End function

	var imageFolder = "<%=ImageDirectory%>/"
	var previousImage
	var selectedImage
	var selectedImageEncoded
	function doSelect(oImage) {
		selectedImage = imageFolder + oImage.childNodes(0).name
		selectedImageEncoded = oImage.childNodes(0).name2
		
		oImage.style.border = "3px solid #08246B"
		currentImage = oImage
		if (previousImage != null) {
			if (previousImage != currentImage) {
				previousImage.style.border = "3px solid #FFFFFF"
			}
		}
		previousImage = currentImage

		document.getElementById("previewWindow").innerHTML = "<img src=" + selectedImage.replace(/ /g, "%20") + ">"
		previewButton.disabled = false
		insertButton.disabled = false

		if (document.getElementById("backgdButton") != null) {
			backgdButton.disabled = false
		}

		if (document.getElementById("deleteButton") != null) {
			deleteButton.disabled = false
		}
	}

	function printStyleList() {
		if (window.opener.document.getElementById("sStyles") != null) {
			document.write(window.opener.document.getElementById("sStyles").outerHTML);
			document.getElementById("sStyles").className = "text70";
				if (document.getElementById("sStyles").options[0].text == "Style") {
					document.getElementById("sStyles").options[0] = null;
					document.getElementById("sStyles").options[0].text = "";
				} else {
					document.getElementById("sStyles").options[1].text = "";
				}

			document.getElementById("sStyles").onchange = null;  
			document.getElementById("sStyles").onmouseenter = null; 
		} else {
			document.write("<select id=sStyles class=text70><option selected></option></select>")
		}
	}

	function printAlign() {
		if ((imageAlign != undefined) && (imageAlign != "")) {
			document.write('<option selected>' + imageAlign)
			document.write('<option>')
		} else {
			document.write('<option selected>')
		}
	}

	function CheckImageForm()
	{
		//upload, upload1, upload2, upload3, upload4
		var imgDir = '<%=ImageDirectory%>';
		var u1 = document.getElementById("upload");
		var u2 = document.getElementById("upload1");
		var u3 = document.getElementById("upload2");
		var u4 = document.getElementById("upload3")
		var u5 = document.getElementById("upload4");

		// Extract just the filename from the paths of the files being uploaded
		u1_file = u1.value;
		last = u1_file.lastIndexOf ("\\", u1_file.length-1);
		u1_file = u1_file.substring (last + 1);

		u2_file = u2.value;
		last = u2_file.lastIndexOf ("\\", u2_file.length-1);
		u2_file = u2_file.substring (last + 1);

		u3_file = u3.value;
		last = u3_file.lastIndexOf ("\\", u3_file.length-1);
		u3_file = u3_file.substring (last + 1);

		u4_file = u4.value;
		last = u4_file.lastIndexOf ("\\", u4_file.length-1);
		u4_file = u4_file.substring (last + 1);

		u5_file = u5.value;
		last = u5_file.lastIndexOf ("\\", u5_file.length-1);
		u5_file = u5_file.substring (last + 1);

		if(u1_file == "" && u2_file == "" && u3_file == "" && u4_file == "" && u5_file == "")
		{
			alert('<%=sTxtChooseImage%>');
			return false;
		}

		// Loop through the imageDir array
		if(u1_file != "")
		{
			for(i = 0; i < imageDir.length; i++)
			{
				if(u1_file == imageDir[i])
				{
					if(!confirm(u1_file + ' <%=sTxtImageExists%>'))
					{
						return false;
					}
				}
			}
		}

		if(u2_file != "")
		{
			for(i = 0; i < imageDir.length; i++)
			{
				if(u2_file == imageDir[i])
				{
					if(!confirm(u2_file + ' <%=sTxtImageExists%>'))
					{
						return false;
					}
				}
			}
		}

		if(u3_file != "")
		{
			for(i = 0; i < imageDir.length; i++)
			{
				if(u3_file == imageDir[i])
				{
					if(!confirm(u3_file + ' <%=sTxtImageExists%>'))
					{
						return false;
					}
				}
			}
		}

		if(u4_file != "")
		{
			for(i = 0; i < imageDir.length; i++)
			{
				if(u4_file == imageDir[i])
				{
					if(!confirm(u4_file + ' <%=sTxtImageExists%>'))
					{
						return false;
					}
				}
			}
		}

		if(u5_file != "")
		{
			for(i = 0; i < imageDir.length; i++)
			{
				if(u5_file == imageDir[i])
				{
					if(!confirm(u5_file + ' <%=sTxtImageExists%>'))
					{
						return false;
					}
				}
			}
		}

		return true;
	}
	</script>

	<title><%=sTxtInsertImage%></title>
	<link rel="stylesheet" href="de_includes/de_styles.css" type="text/css">
	<body bgcolor=threedface style="border: 1px buttonhighlight;">
	<div class="appOutside">
	<div style="border: solid 1px #000000; background-color: #FFFFEE; padding:5px;">
		<img src="de_images/popups/bulb.gif" align=left width=16 height=17>
		<span><%=sTxtInsertImageInst%></span>
	</div>
	<br>

	<form enctype="multipart/form-data" action="<%=HTTPStr%>://<%=serverName%><%=scriptDir%>class.devedit.asp?ToDo=UploadImage&imgDir=<%=ImageDirectory%>&wi=<%=HideWebImage%>&tn=<%=tn%>&dd=<%=dd%>&dt=<%=dt%>&du=<%=du%>" method="post" onSubmit="return CheckImageForm()">
	<!-- Hidden DEP tag is used to persist the path via JS and window.opener -->
		<input type="hidden" name="DEP" value="">
	<!-- End DEP tag -->
	<span class="appInside1" style="width:350px">
		<div class="appInside2">
		<% if du <> "1" then %>
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtUploadImage%></span>
				<br>
				<input type="file" name="upload" class="Text230"> <input type="submit" value="Upload" class="Text75"> <input type="button" value="»" class="Text15" onClick="toggleUploadDiv()" id="toggleButton">
				<span class="err" style="position:absolute; left:40; top:86;"><%=statusText%></span>
				<div id="uploadDiv" style="display:none">
					<input type="file" name="upload1" class="Text230"><br>
					<input type="file" name="upload2" class="Text230"><br>
					<input type="file" name="upload3" class="Text230"><br>
					<input type="file" name="upload4" class="Text230">
				</div>
		<% else %>
			<div class="appInside3" style="padding:11px"><span class="appTitle"><font color="gray"><%=sTxtUploadImage%></font></span>
				<br>
				<input type="file" name="upload" class="Text240" disabled> <input type="submit" value="Upload" class="Text75" disabled>
		<% end if %>
			</div>
		</div>
	</span>
	&nbsp;
	 <% if HideWebImage <> "1" then %>
		<span class="appInside1" style="width:350px">
			<div class="appInside2">
				<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtExternalImage%></span>
					<br>
					<input type="text" name="externalImage" id="externalImage" class="Text240" value="http://">&nbsp;<input type=button value=<%=sTxtLoad%> class="Text75" onClick="insertExtImage()"><br>
					<div style="height:100; display:none" id="dummyDiv">
						&nbsp;
					</div>
				</div>
			</div>
		</span>
	<% else %>
		<span class="appInside1" style="width:350px">
			<div class="appInside2">
				<div class="appInside3" style="padding:11px"><span class="appTitle"><font color="gray"><%=sTxtExternalImage%></font></span>
					<br>
					<input type="text" name="externalImage" id="externalImage" class="Text240" value="http://" disabled>&nbsp;<input type=button value=<%=sTxtLoad%> class="Text75" onClick="insertExtImage()" disabled>
				</div>
			</div>
		</span>
	<% end if %>
	</form>

	<span class="appInside1" style="width:350px">
		<div class="appInside2">
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtInternalImage%></span>
				<table border=0 cellspacing=0 cellpadding=0 style="padding-bottom:5px">
				<tr><td><select style="width:242px; font-size:11px; font-family:Arial;" name="selImageLib">
					<script>outputImageLibraryOptions();</script>
				</select>
				</td><td><input type=button value="<%=sTxtSwitch%>" class=text75 onClick="switchImageLibrary(selImageLib.value)"></td></tr>
				</table>
		<div style="height:325px; width:325px; overflow: auto; border: 2px inset; background-color: #FFFFFF" id="divList">
		<% if Request.QueryString("tn") = 1 then %>
			<table border="0" cellspacing="0" cellpadding="5" style="width:100%">
		<% else %>
			<table border="0" cellspacing="0" cellpadding="3" style="width:100%">
		<% end if %>
			  <tr>
			<%

				if Request.QueryString("tn") = 1 then
					
					for each imageFile in objImageList

					'Workout the image extension
					imageExt = lcase(mid(imageFile.name, instr(imageFile.name, ".")+1, 5))
					isValidExt = false

					for each validImageExt in validImageExts
						if imageExt = validImageExt then
							isValidExt = true
						end if
					next

					'The image extension is valid, lets output it!
					if isValidExt = true then
						counter = counter + 1
					%>
					<td width="25%">
						<span class="body">&nbsp;<%=imageFile.name%><br></span>
						<div onclick="doSelect(this)" style="border: 3px solid #FFFFFF; background-color:#FFFFFF" style="width:80px">
						<img src="<%=ImageDirectory & "/" & imageFile.name %>" width="80" height="80" border=1 name="<%=imageFile.name%>" name2="<%=Server.URLEncode(imageFile.name)%>"></div>
					</td>
					<%
					end if
					
					if counter MOD 3 = 0 then
						response.write "</tr><tr>"
					end if
				
				next
			%>
			
			<% else
			
				for each imageFile in objImageList
					
					'Workout the image extension
					imageExt = lcase(mid(imageFile.name, instr(imageFile.name, ".")+1, 5))
					isValidExt = false

					for each validImageExt in validImageExts
						if imageExt = validImageExt then
							isValidExt = true
						end if
					next

					'The image extension is valid, lets output it!
					if isValidExt = true then
						counter = counter + 1
				%>
					<tr style="cursor:hand">
						<td width="40%" class="body" >
							<div onClick=doSelect(this) style="border: solid 3px #FFFFFF">
							<img src="de_images/popups/image.gif" width=16 height=16 align=absmiddle name="<%=imageFile.name%>" name2="<%=Server.URLEncode(imageFile.name)%>">&nbsp;<%=imageFile.name%>
							<span style="position:absolute; left=200"><%=imageFile.size%> <%=sTxtBytes%></span>
							</div>
						</td>
					</tr>
				<%

					End if

				next
			
			end if

			if counter = 0 then
			%>
				<tr>
					<td width="100%" class="body" >
						<font color="gray"><%=sTxtEmptyImageLibrary%></font>
					</td>
				</tr>
			<%
			end if
			%>

			</table>
			</div>
			</div>
		</div>
	</span>
	&nbsp;
	<span class="appInside1" style="width:350px; position:absolute">
		<div class="appInside2">
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtPreview%></span><br>
				<span id="previewWindow" style="padding:10px; height:150px; width:240px; overflow: auto; border: 2px inset; background-color: #FFFFFF">
				</span><input type="button" name="previewButton" value="<%=sTxtPreview%>" class="Text75" onClick="javascript:viewImage(selectedImage)" disabled=true style="position:absolute; left:257px;">
			</div>
		</div>
	</span>

	<span class="appInside1" style="width:350px; padding-top:5px;">
		<div class="appInside2">
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtImageProperties%></span>
			<table border="0" cellspacing="0" cellpadding="5">
			  <tr>
				<td class="body" width="70"><%=sTxtAltText%>:</td>
				<td class="body" width="88">
				  <input type="text" name="alt_tag" size="50" class="Text70">
				</td>
				<td class="body" width="80"><%=sTxtBorder%>:</td>
				<td class="body" width="80">
				<input type="text" name="border" size="3" class="Text70" maxlength="3" value="0">
				</td>
			  </tr>
			  <tr>
				<td class="body"><%=sTxtImageWidth%>:</td>
				<td class="body">
				  <input type="text" name="image_width" size="3" class="Text70" maxlength="3">
			  </td>
				<td class="body"><%=sTxtImageHeight%>:</td>
				<td class="body">
				  <input type="text" name="image_height" size="3" class="Text70" maxlength="3">
				</td>
			  </tr>
			  <tr>
				<td class="body"><%=sTxtHorizontalSpacing%>:</td>
				<td class="body">
				  <input type="text" name="hspace" size="3" class="Text70" maxlength="3">
				</td>
				<td class="body"><%=sTxtVerticalSpacing%>:</td>
				<td class="body">
				  <input type="text" name="vspace" size="3" class="Text70" maxlength="3">
				</td>
			  </tr>
				<tr>
					<td class="body"><%=sTxtAlignment%>:</td>
					<td class="body">
					  <SELECT class=text70 name=align>
						<script>printAlign()</script>
						<option>Baseline
						<option>Top
						<option>Middle
						<option>Bottom
						<option>TextTop
						<option>ABSMiddle
						<option>ABSBottom
						<option>Left
						<option>Right</option>
					  </select>
					</td>
					<td class="body"><%=sTxtStyle%>:</td>
					<td class="body"><script>printStyleList()</script></td>
				</tr>
			</table>
			</div>
		</div>
	</span>

	<div style="padding-top: 6px;">
	<% if dt <> "0" then %>
	<input type="button" name="backgdButton" value="<%=sTxtImageBackgd%>" class="Text75" onClick="javascript:setBackground(selectedImage)" disabled=true>
	<% end if %>

	<% if dd <> "1" then %>
	<input type="button" name="deleteButton" value="<%=sTxtImageDel%>" class="Text75" onClick="javascript:deleteImage(selectedImageEncoded)" disabled>
	<% end if %>
	</div>

	</div>
	<div style="padding-top: 6px; float: right;">
	<input type="button" name="insertButton" value="<%=sTxtImageInsert%>" class="Text75" onClick="javascript:insertImage(selectedImage)" disabled=true>
	<input type="button" name="Submit" value="<%=sTxtCancel%>" class="Text75" onClick="javascript:window.close()">
	</div>

	</table>

	<script defer>

	if (window.opener.imageEdit)
	{
		selectedImage = window.opener.selectedImage.src;
		previewModify();
	}

	</script>

	<%
		
	end function

	private function PageInsertFlash()

		On Error Resume next

		dt = Request("dt")
		tn = Request("tn")
		dd = Request("dd")
		du = Request("du")
		deveditPath = Request("DEP")

		'Get a list of images from the directory
		set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		validFlashTypes = array("application/x-shockwave-flash")

		' Added for v5.0
		validFlashExts = array("swf")
		isValidExt = false

		FlashDirectory = Request("flashDir")
		HideWebFlash = Request("wi")
		URL = Request.ServerVariables("http_host")
		scriptName = DevEditPath & "/class.devedit.asp"

		'Do we need to kill the following slash?
		serverName = URL
		'if right(serverName, 1) = "/" then
			'serverName = left(serverName, len(serverName)-1)
		'end if

		'Workout the location of class.devedit.asp
		scriptDir = strreverse(Request.ServerVariables("path_info"))
		slashPos = instr(1, scriptDir, "/")
		scriptDir = strreverse(mid(scriptDir, slashPos, len(scriptDir)))
		
		if Request("flashSrc") <> "" then
		
			'Delete the flash file
			flashPath = FlashDirectory & "/" & Request("flashSrc")
			deleteSuccess = objFSO.DeleteFile(Server.MapPath(flashPath))

			if Err.Number <> 0 then
				statusText = sTxtCantDelete
			else
				statusText = sTxtFlashDeleted
			end if

		end if
		
		if request.querystring("ToDo") = "UploadFlash" then

			'Get DT from the querystring
			if Instr(Request("QUERY_STRING"), "dt=0") > 0 then
				dt = 0
			else
				dt = 1
			end if

			if Instr(Request("QUERY_STRING"), "wi=0") > 0 then
				HideWebFlash = 0
			else
				HideWebFlash = 1
			end if

			'Upload the image to the images directory
			set objFile = new Loader
			objFile.Initialize

			newFileName = objFile.GetFileName("upload")
			newFileType = objFile.GetContentType("upload")
			newFileData = objFile.GetFileData("upload")
			validFileType = false
			errorText = ""

			'Is this a valid file type?
			for i = 0 to uBound(validFlashTypes)
				if(newFileType = validFlashTypes(i)) then
					validFileType = true
				end if
			next
		
			if validFileType = false then
				'Invalid file type
				statusText = sTxtFlashErr
			else
			
				uploadSuccess = objFile.SaveToFile("upload", Server.MapPath(FlashDirectory & "/" & newFileName))

				if Err.Number <> 0 then
					statusText = sTxtCantUpload
				else
					statusText = newFileName & " " & sTxtUploadSuccess & "!"	
				end if			
			
			end if
		
		end if

		SFlashDirectory = Server.MapPath(FlashDirectory)

		If (objFSO.FolderExists(SFlashDirectory) = true) Then
			set objFlashDir = objFSO.GetFolder(SFlashDirectory)
		else
			response.write sTxtFlashDirNotConfigured
			response.end
		End if

		' set objFlashDir = objFSO.GetFolder(SFlashDirectory)
		set objFlashList = objFlashDir.Files
		counter = 0

		'Get all images into a JavaScript array so that we can workout whether or not
		'uploading an image would overwrite an existing one
		flashJS = "var flashFiles = Array("

		for each flashFile in objFlashList
			flashJS = flashJS & "'" & flashFile.name & "',"
		next

		if right(flashJS, 1) = "," then
			flashJS = mid(flashJS, 1, len(flashJS)-1)
		end if

		flashJS = flashJS & ")"
	%>
	<title><%=sTxtInsertFlash%></title>
	<link rel="stylesheet" href="de_includes/de_styles.css" type="text/css">

	<script defer>
	if (window.opener.flashEdit) {
		selectedFlash = window.opener.selectedFlash
		previewModify()

	}
	</script>

	<script language=JavaScript>
	window.onload = this.focus

	var selectedFlash
	var selectedFlashFile
	var flashAlign
	var flashLoop

	<%=flashJS%>

	if (window.opener.flashEdit) {
		flashAlign = window.opener.selectedFlash.align
		flashLoop = window.opener.selectedFlash.loop
	}

	function outputFlashLibraryOptions()
	{
		document.write(opener.flashLibs);

		// Loop through all of the image libraries and find the selected one
		for(i = 0; i < selFlashLib.options.length; i++)
		{
			if(selFlashLib.options[i].value == "<%=FlashDirectory%>")
			{
				selFlashLib.selectedIndex = i;
				break;
			}
		}
	}

	function switchFlashLibrary(thePath)
	{
		// Change the path of the flash library
		document.location.href = '<%=HTTPStr%>://<%=serverName%><%=scriptDir%>class.devedit.asp?ToDo=InsertFlash&DEP=<%=Request("DEP")%>&flashDir='+thePath+'&dd=<%=Request.QueryString("dd")%>&du=<%=Request.QueryString("du")%>&wi=<%=Request.QueryString("wi")%>&tn=<%=Request.QueryString("tn")%>&dt=<%=Request.QueryString("dt")%>&wi=<%=HideWebFlash%>';
	}

	function printAlign() {
		if ((flashAlign != undefined) && (flashAlign != "")) {
			document.write('<option selected>' + flashAlign)
			document.write('<option>')
		} else {
			document.write('<option selected>')
		}
	}

	function printLoop() {
		if (flashLoop != undefined) {
			document.write('<option value="' + flashLoop + '" selected>' + flashLoop + '</option>')
			document.write('<option value=""></option>')
		}
	}

	var selectedFlashEmbed
	function previewModify() {

		objectTag = /(<(object|\/object)([\s\S]*?)>)/gi
		paramTag = /(<param([\s\S]*?)>)/gi

		code = selectedFlash.outerHTML.replace(objectTag,"")
		code = code.replace(paramTag,"")
		tempFrame.document.write("<html><head></head><body>" + code + "</body></html>")
		tempFrame.document.close()
		selectedFlashEmbed = tempFrame.document.embeds[0]
		selectedFlashFile = selectedFlash.movie

		document.getElementById("previewWindow").innerHTML = "<embed src='" + selectedFlash.movie + "' quality='high' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='236' height='176' bgcolor='#009933' WMODE=transparent></embed>"

		image_width.value = selectedFlash.width
		image_height.value = selectedFlash.height
		hspace.value = selectedFlash.hspace
		vspace.value = selectedFlash.vspace

		insertButton.value = "<%=sTxtImageModify%>"
		document.title = "<%=sTxtModifyFlash%>"
		previewButton.disabled = false
		insertButton.disabled = false

	}

	function deleteFlash(flashSrc)
	{
		var delImg = confirm("<%=sTxtImageDelete%>");

		if (delImg == true) {
			document.location.href = '<%=HTTPStr%>://<%=serverName%><%=scriptDir%>class.devedit.asp?ToDo=DeleteFlash&DEP=<%=Request("DEP")%>&flashDir=<%=FlashDirectory%>&tn=<%=Request.QueryString("tn")%>&dt=<%=Request.QueryString("dt")%>&wi=<%=HideWebFlash%>&flashSrc='+flashSrc;
		}

	}

	function viewImage(flashSrc)
	{
		var sWidth =  screen.availWidth;
		var sHeight = screen.availHeight;
		
		window.open(flashSrc, 'image', 'width=500, height=500,scrollbars=yes,resizable=yes,left='+(sWidth/2-250)+',top='+(sHeight/2-250));
	}

	function grey(tr) {
			tr.className = 'b4';
	}

	function ungrey(tr) {
			tr.className = '';
	}

	function insertImage(flashSrc) {

		var error = 0;

			imageWidth = image_width.value
			imageHeight = image_height.value
			imageHspace = hspace.value
			imageVspace = vspace.value

			if (isNaN(imageWidth) || imageWidth < 0) {
				alert("<%=sTxtFlashWidthErr%>")
				error = 1
				image_width.select()
				image_width.focus()
			} else if (isNaN(imageHeight) || imageHeight < 0) {
				alert("<%=sTxtFlashHeightErr%>")
				error = 1
				image_height.select()
				image_height.focus()
			} else if (isNaN(imageHspace) || imageHspace < 0) {
				alert("<%=sTxtHorizontalSpacingErr%>")
				error = 1
				hspace.select()
				hspace.focus()
			} else if (isNaN(vspace.value) || vspace.value < 0) {
				alert("<%=sTxtVerticalSpacingErr%>")
				error = 1
				vspace.select()
				vspace.focus()
			}

			if (error != 1) {

				var sel = window.opener.foo.document.selection;
				if (sel!=null) {
					var rng = sel.createRange();
					if (rng!=null) {

						// Are we modifying or inserting?
						if (window.opener.flashEdit) {

							if (imageWidth != "") {
								selectedFlash.width = imageWidth
								selectedFlashEmbed.width = imageWidth
							} else {
								selectedFlash.removeAttribute("width")
								selectedFlashEmbed.removeAttribute("width")
							}

							if (imageHeight != "") {
								selectedFlash.height = imageHeight
								selectedFlashEmbed.height = imageHeight
							} else {
								selectedFlash.removeAttribute("height")
								selectedFlashEmbed.removeAttribute("height")
							}


							if (vspace.value != "") {
								selectedFlash.vspace = vspace.value
								selectedFlashEmbed.vspace = vspace.value
							} else {
								selectedFlash.removeAttribute("vspace")
								selectedFlashEmbed.removeAttribute("vspace")
							}

							if (hspace .value != "") {
								selectedFlash.hspace = hspace.value
								selectedFlashEmbed.hspace = vspace.value
							} else {
								selectedFlash.removeAttribute("hspace")
								selectedFlashEmbed.removeAttribute("hspace")
							}

							if (align[align.selectedIndex].text != "") {
								selectedFlash.align = align[align.selectedIndex].text
							} else {
								selectedFlash.removeAttribute("align")
							}


							selectedFlash.movie = flashSrc

							if (loop[loop.selectedIndex].value != "") {
								selectedFlash.loop =  loop[loop.selectedIndex].value
								selectedFlashEmbed.loop =  loop[loop.selectedIndex].value
							} else {
								selectedFlash.removeAttribute("loop")
								selectedFlashEmbed.removeAttribute("loop")
							}

							embedTag = /(<embed([\s\S]*?)>)/gi
							closeEmbedTag = /(<\/embed([\s\S]*?)>)/gi

							originalFlash = selectedFlash.outerHTML

							code = originalFlash.replace(closeEmbedTag, "")
							code = code.replace(embedTag, selectedFlashEmbed.outerHTML + "</embed>")
							selectedFlash.outerHTML = code

							selectedFlash.runtimeStyle.backgroundImage = "url(<%=deveditPath%>/de_images/hidden.gif)"

						} else {

							if (imageWidth != "")
								imageWidth = ' width=' + imageWidth + '" '
							else
								imageWidth = ''

							if (imageHeight != "")
								imageHeight = ' height=' + imageHeight + '" '
							else
								imageHeight = ''

							if (vspace.value != "")
								vSpace = ' vspace=' + vspace.value + '" '
							else
								vSpace = ''

							if (hspace.value != "")
								hSpace = ' hspace=' + hspace.value + '" '
							else
								hSpace = ''

							if (align[align.selectedIndex].text != "")
								falign = ' align="' + align[align.selectedIndex].text + '" '
							else
								falign = ''

							HTMLTextField = 
							'<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"' + imageHeight + imageWidth + vSpace + hSpace + falign + '>' +
							'<param name=movie value="' + flashSrc + '">' +
							'<param name="LOOP" value="' + loop[loop.selectedIndex].value + '">' + 
							'<embed src="' + flashSrc +
							'" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"' 
							+ imageWidth + imageHeight + vSpace + hSpace + falign + ' loop="' + loop[loop.selectedIndex].value + '"></embed></object>'

							rng.pasteHTML(HTMLTextField)
						}	
					

						// window.opener.foo.focus();
						self.close();

						// oFlash.removeAttribute("id")


					} // End if
				} // End If
			}
	} // End function

	function insertExtFlash() {
		selectedFlashFile = document.getElementById("externalFlash").value
		
		if (previousFlash != null) {
			previousFlash.style.border = "3px solid #FFFFFF"
		}

		document.getElementById("previewWindow").innerHTML = "<embed src='" + selectedFlashFile + "' quality='high' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='236' height='176' bgcolor='#009933' WMODE=transparent></embed>"

		if (document.getElementById("deleteButton") != null) {
			deleteButton.disabled = true
		}

		previewButton.disabled = false
		insertButton.disabled = false

	} // End function

	var flashFolder = "<%=FlashDirectory%>/"
	var previousFlash
	var selectedFlashEncoded
	function doSelect(oFlash) {
		selectedFlashFile = flashFolder + oFlash.childNodes(0).name
		selectedFlashEncoded = oFlash.childNodes(0).name2
		
		oFlash.style.border = "3px solid #08246B"
		currentFlash = oFlash
		if (previousFlash != null) {
			if (previousFlash != currentFlash) {
				previousFlash.style.border = "3px solid #FFFFFF"
			}
		}
		previousFlash = currentFlash

		document.getElementById("previewWindow").innerHTML = "<embed src='" + selectedFlashFile + "' quality='high' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width='236' height='176' bgcolor='#009933' WMODE=transparent></embed>"

		previewButton.disabled = false
		insertButton.disabled = false

		if (document.getElementById("deleteButton") != null) {
			deleteButton.disabled = false
		}
	}

	function CheckFlashForm()
	{
		//upload, upload1, upload2, upload3, upload4
		var flashDir = '<%=FlashDirectory%>';
		var f1 = document.getElementById("upload");

		// Extract just the filename from the paths of the files being uploaded
		f1_file = f1.value;
		last = f1_file.lastIndexOf ("\\", f1_file.length-1);
		f1_file = f1_file.substring (last + 1);

		if(f1_file == "")
		{
			alert('<%=sTxtChooseFlash%>');
			return false;
		}

		// Loop through the imageDir array
		if(f1_file != "")
		{
			for(i = 0; i < flashFiles.length; i++)
			{
				if(f1_file == flashFiles[i])
				{
					if(!confirm(f1_file + ' <%=sTxtFlashExists%>'))
					{
						return false;
					}
				}
			}
		}

		return true;
	}

	</script>
	<title><%=sTxtInsertFlash%></title>

	<body bgcolor=threedface style="border: 1px buttonhighlight;">
	<iframe id=tempFrame style="display:none"></iframe>
	<div class="appOutside">
	<div style="border: solid 1px #000000; background-color: #FFFFEE; padding:5px;">
		<img src="de_images/popups/bulb.gif" align=left width=16 height=17>
		<span><%=sTxtInsertFlashInst%></span>
	</div>
	<br>

	<form enctype="multipart/form-data" action="<%=HTTPStr%>://<%=serverName%><%=scriptDir%>class.devedit.asp?ToDo=UploadFlash&DEP=<%=Request("DEP")%>&flashDir=<%=FlashDirectory%>&wi=<%=HideWebFlash%>&tn=<%=tn%>&dd=<%=dd%>&dt=<%=dt%>" method="post" onSubmit="return CheckFlashForm()">
	<span class="appInside1" style="width:350px">
		<div class="appInside2">
		<% if du <> "1" then %>
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtUploadFlash%></span>
				<br>
					<input type="file" name="upload" class="Text240"> <input type="submit" value="<%=sTxtUpload%>" class="Text75">
					<span class="err" style="position:absolute; left:40; top:86;"><%=statusText%></span>
		<% else %>
			<div class="appInside3" style="padding:11px"><span class="appTitle"><font color="gray"><%=sTxtUploadFlash%></font></span>
				<br>
					<input type="file" name="upload" class="Text240" disabled> <input type="submit" value="<%=sTxtUpload%>" class="Text75" disabled>
		<% end if %>
			</div>
		</div>
	</span>
	&nbsp;
	 <% if HideWebFlash <> "1" then %>
	<span class="appInside1" style="width:350px">
		<div class="appInside2">
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtExternalFlash%></span>
				<br>
				<input type="text" name="externalFlash" id="externalFlash" class="Text240" value="http://">&nbsp;<input type=button value=<%=sTxtLoad%> class="Text75" onClick="insertExtFlash()">
			</div>
		</div>
	</span>
	<% else %>
	<span class="appInside1" style="width:350px">
		<div class="appInside2">
			<div class="appInside3" style="padding:11px"><span class="appTitle"><font color="gray"><%=sTxtExternalFlash%></font></span>
				<br>
				<input type="text" name="externalFlash" id="externalFlash" class="Text240" value="http://" disabled>&nbsp;<input type=button value=<%=sTxtLoad%> class="Text75" onClick="insertExtFlash()" disabled>
			</div>
		</div>
	</span>
	<% end if %>
	</form>

	<span class="appInside1" style="width:350px">
		<div class="appInside2">
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtInternalFlash%></span>
				<table border=0 cellspacing=0 cellpadding=0 style="padding-bottom:5px">
				<tr><td><select style="width:242px; font-size:11px; font-family:Arial;" name="selFlashLib">
					<script>outputFlashLibraryOptions();</script>
				</select>
				</td><td><input type=button value="<%=sTxtSwitch%>" class=text75 onClick="switchFlashLibrary(selFlashLib.value)"></td></tr>
				</table>
		<div style="height:325px; width:325px; overflow: auto; border: 2px inset; background-color: #FFFFFF">
		<% if Request.QueryString("tn") = 1 then %>
			<table border="0" cellspacing="0" cellpadding="5" style="width:100%">
		<% else %>
			<table border="0" cellspacing="0" cellpadding="3" style="width:100%">
		<% end if %>
			  <tr>
			<%

			counter = 0

			if Request.QueryString("tn") = 1 then

				for each flashFile in objFlashList

					'Workout the flash extension
					flashExt = lcase(mid(flashFile.name, instr(flashFile.name, ".")+1, 5))
					isValidExt = false

					for each validflashExt in validFlashExts
						if flashExt = validflashExt then
							isValidExt = true
						end if
					next

					'The image extension is valid, lets output it!
					if isValidExt = true then
						counter = counter + 1
				%>
					<td width="25%">
						<span class="body">&nbsp;<%=flashFile.name%><br></span>
						<div onclick="doSelect(this)" style="border: 3px solid #FFFFFF; background-color:#FFFFFF" style="width:80px">
						<img src="de_images/popups/flash.gif" width="80" height="80" border=1 name="<%=flashFile.name%>" name2="<%=Server.URLEncode(flashFile.name)%>"></div>
					</td>
				<%
				
						counter = counter + 1
					end if
					
					if counter MOD 3 = 0 then
						response.write "</tr><tr>"
					end if
				
				next
			%>
			
			<% else
			
				for each flashFile in objFlashList

					'Workout the image extension
					flashExt = lcase(mid(flashFile.name, instr(flashFile.name, ".")+1, 5))
					isValidExt = false

					for each validflashExt in validFlashExts
						if flashExt = validflashExt then
							isValidExt = true
						end if
					next

					'The image extension is valid, lets output it!
					if isValidExt = true then
						counter = counter + 1

				%>
					<tr style="cursor:hand">
						<td width="40%" class="body" >
							<div onClick=doSelect(this) style="border: solid 3px #FFFFFF">
							<img src="de_images/popups/flash_icon.gif" width=16 height=16 align=absmiddle name="<%=flashFile.name%>" name2="<%=Server.URLEncode(flashFile.name)%>">&nbsp;<%=flashFile.name%>
							<span style="position:absolute; left=200"><%=flashFile.size%> <%=sTxtBytes%></span>
							</div>
						</td>
					</tr>
				<%

					End if

				next
			
			end if
			
			if counter = 0 then
			%>
				<tr>
					<td width="100%" class="body" >
						<font color="gray"><%=sTxtEmptyFlashLibrary%></font>
					</td>
				</tr>
			<%
			end if
			%>

			</table>
			</div>
			</div>
		</div>
	</span>
	&nbsp;
	<span class="appInside1" style="width:350px; position:absolute">
		<div class="appInside2">
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtPreview%></span><br>
				<span id="previewWindow" style="height:180px; width:240px; overflow: auto; border: 2px inset; background-color: #FFFFFF">
				</span><input type="button" name="previewButton" value="<%=sTxtPreview%>" class="Text75" onClick="javascript:viewImage(selectedFlashFile)" disabled=true style="position:absolute; left:257px;">
			</div>
		</div>
	</span>

	<span class="appInside1" style="width:350px; padding-top:5px;">
		<div class="appInside2">
			<div class="appInside3" style="padding:11px"><span class="appTitle"><%=sTxtFlashProperties%></span>
			<table border="0" cellspacing="0" cellpadding="5">
			  <tr>
				<td class="body" width="70"><%=sTxtLoop%>:</td>
				<td class="body" width="88">
					<select class="Text70" name=loop>
						<script>printLoop()</script>
						<option value="true">True</option>
						<option value="false">False</option>
					</select>
				</td>
				<td class="body"><%=sTxtAlignment%>:</td>
					<td class="body">
					  <SELECT class=text70 name=align>
						<script>printAlign()</script>
						<option>Baseline
						<option>Top
						<option>Middle
						<option>Bottom
						<option>TextTop
						<option>ABSMiddle
						<option>ABSBottom
						<option>Left
						<option>Right</option>
					  </select>
					</td>
			  </tr>
			  <tr>
				<td class="body"><%=sTxtFlashWidth%>:</td>
				<td class="body">
				  <input type="text" name="image_width" size="3" class="Text70" maxlength="3">
			  </td>
				<td class="body"><%=sTxtFlashHeight%>:</td>
				<td class="body">
				  <input type="text" name="image_height" size="3" class="Text70" maxlength="3">
				</td>
			  </tr>
			  <tr>
				<td class="body"><%=sTxtHorizontalSpacing%>:</td>
				<td class="body">
				  <input type="text" name="hspace" size="3" class="Text70" maxlength="3">
				</td>
				<td class="body"><%=sTxtVerticalSpacing%>:</td>
				<td class="body">
				  <input type="text" name="vspace" size="3" class="Text70" maxlength="3">
				</td>
			  </tr>
			</table>
			</div>
		</div>
	</span>

	<div style="padding-top: 6px;">
	<input type="button" name="deleteButton" value="<%=sTxtImageDel%>" class="Text75" onClick="javascript:deleteFlash(selectedFlashEncoded)" <% if dd = "1" then %>style=display:none<% end if %> disabled>
	</div>

	</div>
	<div style="padding-top: 6px; float: right;">
	<input type="button" name="insertButton" value="<%=sTxtImageInsert%>" class="Text75" onClick="javascript:insertImage(selectedFlashFile)" disabled=true>
	<input type="button" name="Submit" value="<%=sTxtCancel%>" class="Text75" onClick="javascript:window.close()">
	</div>

	</table>
	
	<%

	end function
%>


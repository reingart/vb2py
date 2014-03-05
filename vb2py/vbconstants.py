
"""VB Constants"""

# Key Codes
vbKeyLButton = 0x1 # Left mouse button
vbKeyRButton = 0x2 # Right mouse button
vbKeyCancel = 0x3 # CANCEL key
vbKeyMButton = 0x4 # Middle mouse button
vbKeyBack = 0x8 # BACKSPACE key
vbKeyTab = 0x9 # TAB key
vbKeyClear = 0xC # CLEAR key
vbKeyReturn = 0xD # ENTER key
vbKeyShift = 0x10 # SHIFT key
vbKeyControl = 0x11 # CTRL key
vbKeyMenu = 0x12 # MENU key
vbKeyPause = 0x13 # PAUSE key
vbKeyCapital = 0x14 # CAPS LOCK key
vbKeyEscape = 0x1B # ESC key
vbKeySpace = 0x20 # SPACEBAR key
vbKeyPageUp = 0x21 # PAGE UP key
vbKeyPageDown = 0x22 # PAGE DOWN key
vbKeyEnd = 0x23 # END key
vbKeyHome = 0x24 # HOME key
vbKeyLeft = 0x25 # LEFT ARROW key
vbKeyUp = 0x26 # UP ARROW key
vbKeyRight = 0x27 # RIGHT ARROW key
vbKeyDown = 0x28 # DOWN ARROW key
vbKeySelect = 0x29 # SELECT key
vbKeyPrint = 0x2A # PRINT SCREEN key
vbKeyExecute = 0x2B # EXECUTE key
vbKeySnapshot = 0x2C # SNAPSHOT key
vbKeyInsert = 0x2D # INSERT key
vbKeyDelete = 0x2E # DELETE key
vbKeyHelp = 0x2F # HELP key
vbKeyNumlock = 0x90 # NUM LOCK key


# Form Codes
vbModeless = 0 # UserForm is modeless.
vbModal = 1 # UserForm is modal (default).


# Colour Codes
vbBlack = 0x0 # Black
vbRed = 0xFF # Red
vbGreen = 0xFF00 # Green
vbYellow = 0xFFFF # Yellow
vbBlue = 0xFF0000 # Blue
vbMagenta = 0xFF00FF # Magenta
vbCyan = 0xFFFF00 # Cyan
vbWhite = 0xFFFFFF # White


# Dir etc Codes
vbNormal = 0 # Normal (default for Dir and SetAttr)
vbReadOnly = 1 # Read-only
vbHidden = 2 # Hidden
vbSystem = 4 # System file
vbVolume = 8 # Volume label
vbDirectory = 16 # Directory or folder
vbArchive = 32 # File has changed since last backup


# File Attribute Codes
Normal = 0 # Normal file. No attributes are set.
ReadOnly = 1 # Read-only file. Attribute is read/write.
Hidden = 2 # Hidden file. Attribute is read/write.
System = 4 # System file. Attribute is read/write.
Volume = 8 # Disk drive volume label. Attribute is read-only.
Directory = 16 # Folder or directory. Attribute is read-only.
Archive = 32 # File has changed since last backup. Attribute is read/write.
Alias = 64 # Link or shortcut. Attribute is read-only.
Compressed = 128 # Compressed file. Attribute is read-only.


# Miscellaneous Codes
vbCrLf = "\n" # Carriage returnlinefeed combination
vbCr = chr(13) # Carriage return character
vbLf = chr(10) # Linefeed character
vbNewLine = "\n" # Platform-specific new line character; whichever is appropriate for current platform
vbNullChar = chr(0) # Character having value 0
vbNullString = chr(0) # String having value 0 Not the same as a zero-length string (""); used for calling external procedures
vbObjectError = -2147221504 # User-defined error numbers should be greater than this value. For example: Err.Raise Number = vbObjectError + 1000
vbTab = chr(9) # Tab character
vbBack = chr(8) # Backspace character
vbFormFeed = chr(12) # Not useful in Microsoft Windows
vbVerticalTab = chr(11) # Not useful in Microsoft Windows


# MsgBox Codes
vbOKOnly = 0 # OK button only (default)
vbOKCancel = 1 # OK and Cancel buttons
vbAbortRetryIgnore = 2 # Abort, Retry, and Ignore buttons
vbYesNoCancel = 3 # Yes, No, and Cancel buttons
vbYesNo = 4 # Yes and No buttons
vbRetryCancel = 5 # Retry and Cancel buttons
vbCritical = 16 # Critical message
vbQuestion = 32 # Warning query
vbExclamation = 48 # Warning message
vbInformation = 64 # Information message
vbDefaultButton1 = 0 # First button is default (default)
vbDefaultButton2 = 256 # Second button is default
vbDefaultButton3 = 512 # Third button is default
vbDefaultButton4 = 768 # Fourth button is default
vbApplicationModal = 0 # Application modal message box (default)
vbSystemModal = 4096 # System modal message box
vbMsgBoxHelpButton = 16384 # Adds Help button to the message box
VbMsgBoxSetForeground = 65536 # Specifies the message box window as the foreground window
vbMsgBoxRight = 524288 # Text is right aligned
vbMsgBoxRtlReading = 1048576 # Specifies text should appear as right-to-left reading on Hebrew and Arabic systems
vbOK = 1 # OK button pressed
vbCancel = 2 # Cancel button pressed
vbAbort = 3 # Abort button pressed
vbRetry = 4 # Retry button pressed
vbIgnore = 5 # Ignore button pressed
vbYes = 6 # Yes button pressed
vbNo = 7 # No button pressed


# Shell Codes
vbHide = 0 # Window is hidden and focus is passed to the hidden window.
vbNormalFocus = 1 # Window has focus and is restored to its original size and position.
vbMinimizedFocus = 2 # Window is displayed as an icon with focus.
vbMaximizedFocus = 3 # Window is maximized with focus.
vbNormalNoFocus = 4 # Window is restored to its most recent size and position. The currently active window remains active.
vbMinimizedNoFocus = 6 # Window is displayed as an icon. The currently active window remains active.


# Special Folder Codes
WindowsFolder = 0 # The Windows folder contains files installed by the Windows operating system.
SystemFolder = 1 # The System folder contains libraries, fonts, and device drivers.
TemporaryFolder = 2 # The Temp folder is used to store temporary files. Its path is found in the TMP environment variable.


# StrConv Codes
vbUpperCase = 1 # Converts the string to uppercase characters.
vbLowerCase = 2 # Converts the string to lowercase characters.
vbProperCase = 3 # Converts the first letter of every word in string to uppercase.
vbWide = 4 # Converts narrow (single-byte) characters in string to wide (double-byte) characters. Applies to Far East locales.
vbNarrow = 8 # Converts wide (double-byte) characters in string to narrow (single-byte) characters. Applies to Far East locales.
vbKatakana = 16 # Converts Hiragana characters in string to Katakana characters. Applies to Japan only.
vbHiragana = 32 # Converts Katakana characters in string to Hiragana characters. Applies to Japan only.
vbUnicode = 64 # Converts the string to Unicode using the default code page of the system.
vbFromUnicode = 128 # Converts the string from Unicode to the default code page of the system.


# System Colour Codes
vbScrollBars = 0x80000000 # Scroll bar color
vbDesktop = 0x80000001 # Desktop color
vbActiveTitleBar = 0x80000002 # Color of the title bar for the active window
vbInactiveTitleBar = 0x80000003 # Color of the title bar for the inactive window
vbMenuBar = 0x80000004 # Menu background color
vbWindowBackground = 0x80000005 # Window background color
vbWindowFrame = 0x80000006 # Window frame color
vbMenuText = 0x80000007 # Color of text on menus
vbWindowText = 0x80000008 # Color of text in windows
vbTitleBarText = 0x80000009 # Color of text in caption, size box, and scroll arrow
vbActiveBorder = 0x8000000A # Border color of active window
vbInactiveBorder = 0x8000000B # Border color of inactive window
vbApplicationWorkspace = 0x8000000C # Background color of multiple-document interface (MDI) applications
vbHighlight = 0x8000000D # Background color of items selected in a control
vbHighlightText = 0x8000000E # Text color of items selected in a control
vbButtonFace = 0x8000000F # Color of shading on the face of command buttons
vbButtonShadow = 0x80000010 # Color of shading on the edge of command buttons
vbGrayText = 0x80000011 # Grayed (disabled) text
vbButtonText = 0x80000012 # Text color on push buttons
vbInactiveCaptionText = 0x80000013 # Color of text in an inactive caption
vb3DHighlight = 0x80000014 # Highlight color for 3-D display elements
vb3DDKShadow = 0x80000015 # Darkest shadow color for 3-D display elements
vb3DLight = 0x80000016 # Second lightest 3-D color after vb3DHighlight
vbInfoText = 0x80000017 # Color of text in ToolTips
vbInfoBackground = 0x80000018 # Background color of ToolTips


# Var Type Codes
vbEmpty = 0 # Uninitialized (default)
vbNull = 1 # Contains no valid data
vbInteger = 2 # Integer
vbLong = 3 # Long integer
vbSingle = 4 # Single-precision floating-point number
vbDouble = 5 # Double-precision floating-point number
vbCurrency = 6 # Currency
vbDate = 7 # Date
vbString = 8 # String
vbObject = 9 # Object
vbError = 10 # Error
vbBoolean = 11 # Boolean
vbVariant = 12 # Variant (used only for arrays of variants)
vbDataObject = 13 # Data access object
vbDecimal = 14 # Decimal
vbByte = 17 # Byte
vbUserDefinedType = 36 # Variants that contain user-defined types
vbArray = 8192 # Array



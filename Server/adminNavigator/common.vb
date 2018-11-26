
Namespace Contensive.adminNavigator
    Module common
        '
        Public Const iconCloseWindow As String = "<i title=close class=""fas fa-window-close"" style=""color:#fff""></i>"
        Public Const iconOpenWindow As String = "<i title=open class=""fa fa-angle-double-right"" style=""color:#fff""></i>"

        '
        '
        '
        Public Const IconAdvanced = "<img src=""/cclib/images/NavAdv.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconAdvancedClosed = "<img src=""/cclib/images/NavAdvClosed.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconAdvancedOpened = "<img src=""/cclib/images/NavAdvOpened.gif"" class=""ccImgA"" title=""{title}"">"
        'Public Const IconAdvanced = "<img src=""/cclib/images/NavAdv.gif"" width=18 x height=18 border=0 style=""vertical-align:middle;"" title=""{title}"">"
        'Public Const IconAdvancedClosed = "<img src=""/cclib/images/NavAdvClosed.gif"" width=18 x height=18 border=0 style=""vertical-align:middle;"" title=""{title}"">"
        'Public Const IconAdvancedOpened = "<img src=""/cclib/images/NavAdvOpened.gif"" width=18 x height=18 border=0 style=""vertical-align:middle;"" title=""{title}"">"
        '
        Public Const IconContent = "<img src=""/cclib/images/NavContent.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconContentOpened = "<img src=""/cclib/images/NavContentOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconContentClosed = "<img src=""/cclib/images/NavContentClosed.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconFolder = "<img src=""/cclib/images/NavFolderClosed.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconFolderClosed = "<img src=""/cclib/images/NavFolderClosed.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconFolderOpened = "<img src=""/cclib/images/NavFolderOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconFolderNoSubNodes = "<img src=""/cclib/images/NavFolder.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconEmail = "<img src=""/cclib/images/NavEmail.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconEmailClosed = "<img src=""/cclib/images/NavEmailClosed.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconEmailOpened = "<img src=""/cclib/images/NavEmailOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconSystemEmail = "<img src=""/cclib/images/NavEmail.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconConditionalEmail = "<img src=""/cclib/images/NavEmail.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconGroupEmail = "<img src=""/cclib/images/NavEmail.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconUsers = "<img src=""/cclib/images/NavUsers.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconUsersOpened = "<img src=""/cclib/images/NavUsersOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconUsersClosed = "<img src=""/cclib/images/NavUsersClosed.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconReports = "<img src=""/cclib/images/NavReports.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconReportsOpened = "<img src=""/cclib/images/NavReportsOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconReportsClosed = "<img src=""/cclib/images/NavReportsClosed.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconSettings = "<img src=""/cclib/images/NavSettings.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconSettingsOpened = "<img src=""/cclib/images/NavSettingsOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconSettingsClosed = "<img src=""/cclib/images/NavSettingsClosed.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconTools = "<img src=""/cclib/images/NavTools.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconToolsOpened = "<img src=""/cclib/images/NavToolsOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconToolsClosed = "<img src=""/cclib/images/NavToolsClosed.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconRecord = "<img src=""/cclib/images/NavRecord.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconRecordOpened = "<img src=""/cclib/images/NavRecord.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconRecordClosed = "<img src=""/cclib/images/NavRecord.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconHelp = "<img src=""/cclib/images/NavHelp.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconAddon = "<img src=""/cclib/images/NavAddons.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconAddons = "<img src=""/cclib/images/NavAddons.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconAddonsOpened = "<img src=""/cclib/images/NavAddonsOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconAddonsClosed = "<img src=""/cclib/images/NavAddonsClosed.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const IconPublicHome = "<img src=""/cclib/images/NavPublicHome.gif"" class=""ccImgA"" title=""Public Home"">"
        Public Const IconAdminHome = "<img src=""/cclib/images/NavHome.gif"" class=""ccImgA"" title=""Admin Home"">"
        '
        Public Const IconBox = "<img src=""/cclib/mktree/box.gif"" width=19 height=19 border=0 style=""vertical-align:middle;"" title=""{title}"">"
        '
        Public Const IconManageContent = "<img src=""/cclib/images/NavContent.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconManageContentOpened = "<img src=""/cclib/images/NavContentOpened.gif"" class=""ccImgA"" title=""{title}"">"
        Public Const IconManageContentClosed = "<img src=""/cclib/images/NavContentClosed.gif"" class=""ccImgA"" title=""{title}"">"
        '
        Public Const NavigatorContentName = "Navigator Entries"
        Public Const LegacyMenuContentName = "Menu Entries"
        Public Const NodeIDAllContentList = "NodeIDAllContentList"
        Public Const NodeIDAddonsNoCollection = "NodeIDAddonsNoCollection"
        Public Const NodeIDLegacyMenu = "legacymenunode"
        Public Const NodeIDManageAddons = "manageaddons"
        Public Const NodeIDManageAddonsAdvanced = "manageaddonsadvanced"
        Public Const NodeIDManageAddonsCollectionPrefix = "collection"
        Public Const NodeIDSettings = "settings"
        Public Const NodeIDTools = "tools"
        Public Const NodeIDReports = "reports"
        '
        Public Structure SortNodeType
            Public Name As String
            Public addonid As Integer
            Public ContentControlID As Integer
            Public CollectionID As Integer
            Public NavigatorID As Integer
            Public NewWindow As Boolean
            Public ContentID As Integer
            Public Link As String
            Public NavIconType As Integer
            Public NavIconTitle As String
            Public HelpAddonID As Integer
            Public helpCollectionID As Integer
            Public NodeIDString As String
        End Structure
        '
        Public Enum NodeTypeEnum
            NodeTypeEntry = 0
            NodeTypeCollection = 1
            NodeTypeAddon = 2
            NodeTypeContent = 3
        End Enum
        '        '
        '        ' kmaModule
        '        '
        '        '
        '        '========================================================================
        '        '   kma defined errors
        '        '       1000-1999 Contensive
        '        '       2000-2999 Datatree
        '        '
        '        '   see kmaErrorDescription() for transations
        '        '========================================================================
        '        '
        '        Const Error_DataTree_RootNodeNext = 2000
        '        Const Error_DataTree_NoGoNext = 2001
        '        '
        '        '========================================================================
        '        '
        '        '========================================================================
        '        '
        '        Declare Function GetTickCount Lib "kernel32" () as integer
        '        Declare Function GetCurrentProcessId Lib "kernel32" () as integer
        '        Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as integer)
        '        '
        '        '========================================================================
        '        '   Declarations for SetPiorityClass
        '        '========================================================================
        '        '
        '        Const THREAD_BASE_PRIORITY_IDLE = -15
        '        Const THREAD_BASE_PRIORITY_LOWRT = 15
        '        Const THREAD_BASE_PRIORITY_MIN = -2
        '        Const THREAD_BASE_PRIORITY_MAX = 2
        '        Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
        '        Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
        '        Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
        '        Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
        '        Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
        '        Const THREAD_PRIORITY_NORMAL = 0
        '        Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
        '        Const HIGH_PRIORITY_CLASS = &H80
        '        Const IDLE_PRIORITY_CLASS = &H40
        '        Const NORMAL_PRIORITY_CLASS = &H20
        '        Const REALTIME_PRIORITY_CLASS = &H100
        '        '
        '        Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread as integer, ByVal nPriority as integer) as integer
        '        Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess as integer, ByVal dwPriorityClass as integer) as integer
        '        Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread as integer) as integer
        '        Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess as integer) as integer
        '        Private Declare Function GetCurrentThread Lib "kernel32" () as integer
        '        Private Declare Function GetCurrentProcess Lib "kernel32" () as integer
        '        '

        '        '
        '        '========================================================================
        '        'Converts unsafe characters,
        '        'such as spaces, into their
        '        'corresponding escape sequences.
        '        '========================================================================
        '        '
        '        Declare Function UrlEscape Lib "shlwapi" _
        '           Alias "UrlEscapeA" _
        '          (ByVal pszURL As String, _
        '           ByVal pszEscaped As String, _
        '           pcchEscaped as integer, _
        '           ByVal dwFlags as integer) as integer
        '        '
        '        'Converts escape sequences back into
        '        'ordinary characters.
        '        '
        '        Declare Function UrlUnescape Lib "shlwapi" _
        '           Alias "UrlUnescapeA" _
        '          (ByVal pszURL As String, _
        '           ByVal pszUnescaped As String, _
        '           pcchUnescaped as integer, _
        '           ByVal dwFlags as integer) as integer

        '        '
        '        '   Error reporting strategy
        '        '       Popups are pop-up boxes that tell the user what to do
        '        '       Logs are files with error details for developers to use debugging
        '        '
        '        '       Attended Programs
        '        '           - errors that do not effect the operation, resume next
        '        '           - all errors trickle up to the user interface level
        '        '           - User Errors, like file not found, return "UserError" code and a description
        '        '           - Internal Errors, like div-by-0, User should see no detail, log gets everything
        '        '           - Dependant Object Errors, codes that return from objects:
        '        '               - If UserError, translate ErrSource for raise, but log all original info
        '        '               - If InternalError, log info and raise InternalError
        '        '               - If you can not tell, call it InternalError
        '        '
        '        '       UnattendedMode
        '        '           The same, except each routine decides when
        '        '
        '        '       When an error happens in-line (bad condition without a raise)
        '        '           Log the error
        '        '           Raise the appropriate Code/Description in the current Source
        '        '
        '        '       When an ErrorTrap occurs
        '        '           If ErrSource is not AppTitle, it is a dependantObjectError, log and translate code
        '        '           If ErrNumber is not an ObjectError, call it internal error, log and translate code
        '        '           Error must be either "InternalError" or "UserError", just raise it again
        '        '
        '        ' old - If an error is raised that is not a KmaCode, it is logged and translated
        '        ' old - If an error is raised and the soure is not he current App.EXEname, it is logged and translated
        '        '
        '        Public Const KmaErrorBase = vbObjectError                 ' Base on which Internal errors should start
        '        '
        '        'Public Const KmaError_UnderlyingObject = vbObjectError + 1     ' An error occurec in an underlying object
        '        'Public Const KmaccErrorServiceStopped = vbObjectError + 2       ' The service is not running
        '        'Public Const KmaError_BadObject = vbObjectError + 3            ' The Server Pointer is not valid
        '        'Public Const KmaError_UpgradeInProgress = vbObjectError + 4    ' page is blocked because an upgrade is in progress
        '        'Public Const KmaError_InvalidArgument = vbObjectError + 5      ' and input argument is not valid. Put details at end of description
        '        '
        '        Public Const KmaErrorUser = KmaErrorBase + 16                   ' Generic Error code that passes the description back to the user
        '        Public Const KmaErrorInternal = KmaErrorBase + 17               ' Internal error which the user should not see
        '        Public Const KmaErrorPage = KmaErrorBase + 18                   ' Error from the page which called Contensive
        '        '
        '        Public Const KmaObjectError = KmaErrorBase + 256                ' Internal error which the user should not see
        '        '
        '        '==========================================================================
        '        '       NTSvc.ocx, LogEvent Constants
        '        '==========================================================================
        '        '
        '        Public Const NTServiceEventError = 1                    ' Error event.
        '        Public Const NTServiceEventWarning = 2                  ' Warning event.
        '        Public Const NTServiceEventInformation = 4              ' Information event.
        '        Public Const NTServiceEventAuditSuccess = 8             ' Audit success event.
        '        Public Const NTServiceEventAuditFailure = 10            ' Audit failure event.

        '        Public Const NTServiceIDDebug = 108                     ' Debugging message
        '        Public Const NTServiceIDError = 109                     ' Error message
        '        Public Const NTServiceIDInfo = 110                      ' Information message
        '        '
        '        '==========================================================================
        '        '       NTSvc.ocx, LogEvent Constants
        '        '==========================================================================
        '        '
        '        Public Const SQLTrue = "1"
        '        Public Const SQLFalse = "0"
        '        '
        '        '
        '        '
        '        Public Const kmaEndTable = "</table >"
        '        Public Const kmaEndTableCell = "</td>"
        '        Public Const kmaEndTableRow = "</tr>"
        '        '
        '        '==========================================================================
        '        ' kmaByteArrayToString / kmaStringToByteArray
        '        '==========================================================================
        '        '
        '        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy as integer)
        '        'The WideCharToMultiByte function maps a wide-character string to a new character string.
        '        'The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.
        '        'CodePage
        '        Private Const CP_ACP = 0 'ANSI
        '        Private Const CP_MACCP = 2 'Mac
        '        Private Const CP_OEMCP = 1 'OEM
        '        Private Const CP_UTF7 = 65000
        '        Private Const CP_UTF8 = 65001
        '        'dwFlags
        '        Private Const WC_NO_BEST_FIT_CHARS = &H400
        '        Private Const WC_COMPOSITECHECK = &H200
        '        Private Const WC_DISCARDNS = &H10
        '        Private Const WC_SEPCHARS = &H20 'Default
        '        Private Const WC_DEFAULTCHAR = &H40

        '        Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
        '            ByVal CodePage as integer, _
        '            ByVal dwFlags as integer, _
        '            ByVal lpWideCharStr as integer, _
        '            ByVal cchWideChar as integer, _
        '            ByVal lpMultiByteStr as integer, _
        '            ByVal cbMultiByte as integer, _
        '            ByVal lpDefaultChar as integer, _
        '            ByVal lpUsedDefaultChar as integer) as integer
        '
        '==========================================================================
        '   Convert a variant to an long (long)
        '   returns 0 if the input is not an integer
        '   if float, rounds to integer
        '==========================================================================
        '
        Public Function kmaEncodeInteger(ExpressionVariant As Object) As Integer
            Dim test As String = ExpressionVariant.ToString

            kmaEncodeInteger = 0
            If IsNumeric(test) Then
                kmaEncodeInteger = CInt(test)
            End If
            '
            Exit Function
            '
            ' ----- ErrorTrap
            '
ErrorTrap:
            Err.Clear()
            kmaEncodeInteger = 0
        End Function
        '        '
        '        '==========================================================================
        '        '   Convert a variant to a number (double)
        '        '   returns 0 if the input is not a number
        '        '==========================================================================
        '        '
        '        Public Function KmaEncodeNumber(ExpressionVariant As Object) As Double
        '            On Error GoTo ErrorTrap
        '            '
        '            'KmaEncodeNumber = 0
        '            If Not IsMissing(ExpressionVariant) Then
        '                If Not IsNull(ExpressionVariant) Then
        '                    If ExpressionVariant <> "" Then
        '                        If IsNumeric(ExpressionVariant) Then
        '                            KmaEncodeNumber = ExpressionVariant
        '                        End If
        '                    End If
        '                End If
        '            End If
        '            '
        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '            KmaEncodeNumber = 0
        '        End Function
        '        '
        '        '==========================================================================
        '        '   Convert a variant to a date
        '        '   returns 0 if the input is not a number
        '        '==========================================================================
        '        '
        '        Public Function KmaEncodeDate(ExpressionVariant As Object) As Date
        '            On Error GoTo ErrorTrap
        '            '
        '            '    KmaEncodeDate = CDate(ExpressionVariant)
        '            '    KmaEncodeDate = CDate("1/1/1980")
        '            'KmaEncodeDate = CDate(0)
        '            If Not IsMissing(ExpressionVariant) Then
        '                If Not IsNull(ExpressionVariant) Then
        '                    If ExpressionVariant <> "" Then
        '                        If IsDate(ExpressionVariant) Then
        '                            KmaEncodeDate = ExpressionVariant
        '                        End If
        '                    End If
        '                End If
        '            End If
        '            '
        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '            KmaEncodeDate = CDate(0)
        '        End Function
        '        '
        '        '==========================================================================
        '        '   Convert a variant to a boolean
        '        '   Returns true if input is not false, else false
        '        '==========================================================================
        '        '
        '        Public Function kmaEncodeBoolean(ExpressionVariant As Object) As Boolean
        '            On Error GoTo ErrorTrap
        '            '
        '            'KmaEncodeBoolean = False
        '            If Not IsMissing(ExpressionVariant) Then
        '                If Not IsNull(ExpressionVariant) Then
        '                    If ExpressionVariant <> "" Then
        '                        If IsNumeric(ExpressionVariant) Then
        '                            If ExpressionVariant <> "0" Then
        '                                If ExpressionVariant <> 0 Then
        '                                    kmaEncodeBoolean = True
        '                                End If
        '                            End If
        '                        ElseIf UCase(ExpressionVariant) = "ON" Then
        '                            kmaEncodeBoolean = True
        '                        ElseIf UCase(ExpressionVariant) = "YES" Then
        '                            kmaEncodeBoolean = True
        '                        ElseIf UCase(ExpressionVariant) = "TRUE" Then
        '                            kmaEncodeBoolean = True
        '                        Else
        '                            kmaEncodeBoolean = False
        '                        End If
        '                    End If
        '                End If
        '            End If
        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '            kmaEncodeBoolean = False
        '        End Function
        '        '
        '        '==========================================================================
        '        '   Convert a variant into 0 or 1
        '        '   Returns 1 if input is not false, else 0
        '        '==========================================================================
        '        '
        '        Public Function KmaEncodeBit(ExpressionVariant As Object) as integer
        '            On Error GoTo ErrorTrap
        '            '
        '            'KmaEncodeBit = 0
        '            If kmaEncodeBoolean(ExpressionVariant) Then
        '                KmaEncodeBit = 1
        '            End If
        '            '
        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '            KmaEncodeBit = 0
        '        End Function
        '        '
        '        '==========================================================================
        '        '   Convert a variant to a string
        '        '   returns emptystring if the input is not a string
        '        '==========================================================================
        '        '
        '        Public Function kmaEncodeText(ExpressionVariant As Object) As String
        '            On Error GoTo ErrorTrap
        '            '
        '            'KmaEncodeText = ""
        '            If Not IsMissing(ExpressionVariant) Then
        '                If Not IsNull(ExpressionVariant) Then
        '                    kmaEncodeText = CStr(ExpressionVariant)
        '                End If
        '            End If
        '            '
        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '            kmaEncodeText = ""
        '        End Function
        '        '
        '        '==========================================================================
        '        '   Converts a possibly missing value to variant
        '        '==========================================================================
        '        '
        '        Public Function KmaEncodeMissing(ExpressionVariant As Object, DefaultVariant As Object) As Object
        '            'On Error GoTo ErrorTrap
        '            '
        '            If IsMissing(ExpressionVariant) Then
        '                KmaEncodeMissing = DefaultVariant
        '            Else
        '                KmaEncodeMissing = ExpressionVariant
        '            End If
        '            '
        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function KmaEncodeMissingText(ExpressionVariant As Object, DefaultText As String) As String
        '            KmaEncodeMissingText = kmaEncodeText(KmaEncodeMissing(ExpressionVariant, DefaultText))
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function KmaEncodeMissingInteger(ExpressionVariant As Object, DefaultInteger as integer) as integer
        '            KmaEncodeMissingInteger = cp.Utils.EncodeInteger(KmaEncodeMissing(ExpressionVariant, DefaultInteger))
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function KmaEncodeMissingDate(ExpressionVariant As Object, DefaultDate As Date) As Date
        '            KmaEncodeMissingDate = KmaEncodeDate(KmaEncodeMissing(ExpressionVariant, DefaultDate))
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function KmaEncodeMissingNumber(ExpressionVariant As Object, DefaultNumber As Double) As Double
        '            KmaEncodeMissingNumber = KmaEncodeNumber(KmaEncodeMissing(ExpressionVariant, DefaultNumber))
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function KmaEncodeMissingBoolean(ExpressionVariant As Object, DefaultState As Boolean) As Boolean
        '            KmaEncodeMissingBoolean = kmaEncodeBoolean(KmaEncodeMissing(ExpressionVariant, DefaultState))
        '        End Function
        '        '
        '        '================================================================================================================
        '        '   Separate a URL into its host, path, page parts
        '        '================================================================================================================
        '        '
        '        Public Sub SeparateURL(ByVal SourceURL As String, ByRef Protocol As String, ByRef Host As String, ByRef Path As String, ByRef Page As String, ByRef QueryString As String)
        '            'On Error GoTo ErrorTrap
        '            '
        '            '   Divide the URL into URLHost, URLPath, and URLPage
        '            '
        '            Dim WorkingURL As String
        '            Dim Position as integer
        '            '
        '            ' Get Protocol (before the first :)
        '            '
        '            WorkingURL = SourceURL
        '            Position = InStr(1, WorkingURL, ":")
        '            'Position = InStr(1, WorkingURL, "://")
        '            If Position <> 0 Then
        '                Protocol = Mid(WorkingURL, 1, Position + 2)
        '                WorkingURL = Mid(WorkingURL, Position + 3)
        '            End If
        '            '
        '            ' compatibility fix
        '            '
        '            If InStr(1, WorkingURL, "//") = 1 Then
        '                If Protocol = "" Then
        '                    Protocol = "http:"
        '                End If
        '                Protocol = Protocol & "//"
        '                WorkingURL = Mid(WorkingURL, 3)
        '            End If
        '            '
        '            ' Get QueryString
        '            '
        '            Position = InStr(1, WorkingURL, "?")
        '            If Position > 0 Then
        '                QueryString = Mid(WorkingURL, Position)
        '                WorkingURL = Mid(WorkingURL, 1, Position - 1)
        '            End If
        '            '
        '            ' separate host from pathpage
        '            '
        '            'iURLHost = WorkingURL
        '            Position = InStr(WorkingURL, "/")
        '            If (Position = 0) And (Protocol = "") Then
        '                '
        '                ' Page without path or host
        '                '
        '                Page = WorkingURL
        '                Path = ""
        '                Host = ""
        '            ElseIf (Position = 0) Then
        '                '
        '                ' host, without path or page
        '                '
        '                Page = ""
        '                Path = "/"
        '                Host = WorkingURL
        '            Else
        '                '
        '                ' host with a path (at least)
        '                '
        '                Path = Mid(WorkingURL, Position)
        '                Host = Mid(WorkingURL, 1, Position - 1)
        '                '
        '                ' separate page from path
        '                '
        '                Position = InStrRev(Path, "/")
        '                If Position = 0 Then
        '                    '
        '                    ' no path, just a page
        '                    '
        '                    Page = Path
        '                    Path = "/"
        '                Else
        '                    Page = Mid(Path, Position + 1)
        '                    Path = Mid(Path, 1, Position)
        '                End If
        '            End If
        '            Exit Sub
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '        End Sub
        '        '
        '        '================================================================================================================
        '        '   Separate a URL into its host, path, page parts
        '        '================================================================================================================
        '        '
        '        Public Sub ParseURL(ByVal SourceURL As String, ByRef Protocol As String, ByRef Host As String, ByRef Port As String, ByRef Path As String, ByRef Page As String, ByRef QueryString As String)
        '            'On Error GoTo ErrorTrap
        '            '
        '            '   Divide the URL into URLHost, URLPath, and URLPage
        '            '
        '            Dim iURLWorking As String               ' internal storage for GetURL functions
        '            Dim iURLProtocol As String
        '            Dim iURLHost As String
        '            Dim iURLPort As String
        '            Dim iURLPath As String
        '            Dim iURLPage As String
        '            Dim iURLQueryString As String
        '            Dim Position as integer
        '            '
        '            iURLWorking = SourceURL
        '            Position = InStr(1, iURLWorking, "://")
        '            If Position <> 0 Then
        '                iURLProtocol = Mid(iURLWorking, 1, Position + 2)
        '                iURLWorking = Mid(iURLWorking, Position + 3)
        '            End If
        '            '
        '            ' separate Host:Port from pathpage
        '            '
        '            iURLHost = iURLWorking
        '            Position = InStr(iURLHost, "/")
        '            If Position = 0 Then
        '                '
        '                ' just host, no path or page
        '                '
        '                iURLPath = "/"
        '                iURLPage = ""
        '            Else
        '                iURLPath = Mid(iURLHost, Position)
        '                iURLHost = Mid(iURLHost, 1, Position - 1)
        '                '
        '                ' separate page from path
        '                '
        '                Position = InStrRev(iURLPath, "/")
        '                If Position = 0 Then
        '                    '
        '                    ' no path, just a page
        '                    '
        '                    iURLPage = iURLPath
        '                    iURLPath = "/"
        '                Else
        '                    iURLPage = Mid(iURLPath, Position + 1)
        '                    iURLPath = Mid(iURLPath, 1, Position)
        '                End If
        '            End If
        '            '
        '            ' Divide Host from Port
        '            '
        '            Position = InStr(iURLHost, ":")
        '            If Position = 0 Then
        '                '
        '                ' host not given, take a guess
        '                '
        '                Select Case UCase(iURLProtocol)
        '                    Case "FTP://"
        '                        iURLPort = "21"
        '                    Case "HTTP://", "HTTPS://"
        '                        iURLPort = "80"
        '                    Case Else
        '                        iURLPort = "80"
        '                End Select
        '            Else
        '                iURLPort = Mid(iURLHost, Position + 1)
        '                iURLHost = Mid(iURLHost, 1, Position - 1)
        '            End If
        '            Position = InStr(1, iURLPage, "?")
        '            If Position > 0 Then
        '                iURLQueryString = Mid(iURLPage, Position)
        '                iURLPage = Mid(iURLPage, 1, Position - 1)
        '            End If
        '            Protocol = iURLProtocol
        '            Host = iURLHost
        '            Port = iURLPort
        '            Path = iURLPath
        '            Page = iURLPage
        '            QueryString = iURLQueryString
        '            Exit Sub
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '        End Sub
        '        '
        '        '
        '        '
        '        Function DecodeGMTDate(GMTDate As String) As Date
        '            'On Error GoTo ErrorTrap
        '            '
        '            Dim WorkString As String
        '            DecodeGMTDate = 0
        '            If GMTDate <> "" Then
        '                WorkString = Mid(GMTDate, 6, 11)
        '                If IsDate(WorkString) Then
        '                    DecodeGMTDate = CDate(WorkString)
        '                    WorkString = Mid(GMTDate, 18, 8)
        '                    If IsDate(WorkString) Then
        '                        DecodeGMTDate = DecodeGMTDate + CDate(WorkString) + 4 / 24
        '                    End If
        '                End If
        '            End If
        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '        End Function
        '        '
        '        '
        '        '
        '        Function EncodeGMTDate(MSDate As Date) As String
        '            'On Error GoTo ErrorTrap
        '            '
        '            Dim WorkString As String
        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '        End Function
        '        '
        '        '=================================================================================
        '        '   Renamed to catch all the cases that used it in addons
        '        '
        '        '   Do not use this routine in Addons to get the addon option string value
        '        '   to get the value in an option string, use csv.getAddonOption("name")
        '        '
        '        ' Get the value of a name in a string of name value pairs parsed with vrlf and =
        '        '   the legacy line delimiter was a '&' -> name1=value1&name2=value2"
        '        '   new format is "name1=value1 crlf name2=value2 crlf ..."
        '        '   There can be no extra spaces between the delimiter, the name and the "="
        '        '=================================================================================
        '        '
        '        Function getSimpleNameValue(Name As String, ArgumentString As String, DefaultValue As String, Delimiter As String) As String
        '            'Function getArgument(Name As String, ArgumentString As String, Optional DefaultValue As Variant, Optional Delimiter As String) As String
        '            '
        '            Dim WorkingString As String
        '            Dim iDefaultValue As String
        '            Dim NameLength as integer
        '            Dim ValueStart as integer
        '            Dim ValueEnd as integer
        '            Dim IsQuoted As Boolean
        '            '
        '            ' determine delimiter
        '            '
        '            If Delimiter = "" Then
        '                '
        '                ' If not explicit
        '                '
        '                If InStr(1, ArgumentString, vbCrLf) <> 0 Then
        '                    '
        '                    ' crlf can only be here if it is the delimiter
        '                    '
        '                    Delimiter = vbCrLf
        '                Else
        '                    '
        '                    ' either only one option, or it is the legacy '&' delimit
        '                    '
        '                    Delimiter = "&"
        '                End If
        '            End If
        '            iDefaultValue = KmaEncodeMissing(DefaultValue, "")
        '            WorkingString = ArgumentString
        '            getSimpleNameValue = iDefaultValue
        '            If WorkingString <> "" Then
        '                WorkingString = Delimiter & WorkingString & Delimiter
        '                ValueStart = InStr(1, WorkingString, Delimiter & Name & "=", vbTextCompare)
        '                If ValueStart <> 0 Then
        '                    NameLength = Len(Name)
        '                    ValueStart = ValueStart + Len(Delimiter) + NameLength + 1
        '                    If Mid(WorkingString, ValueStart, 1) = """" Then
        '                        IsQuoted = True
        '                        ValueStart = ValueStart + 1
        '                    End If
        '                    If IsQuoted Then
        '                        ValueEnd = InStr(ValueStart, WorkingString, """" & Delimiter)
        '                    Else
        '                        ValueEnd = InStr(ValueStart, WorkingString, Delimiter)
        '                    End If
        '                    If ValueEnd = 0 Then
        '                        getSimpleNameValue = Mid(WorkingString, ValueStart)
        '                    Else
        '                        getSimpleNameValue = Mid(WorkingString, ValueStart, ValueEnd - ValueStart)
        '                    End If
        '                End If
        '            End If
        '            '

        '            Exit Function
        '            '
        '            ' ----- ErrorTrap
        '            '
        'ErrorTrap:
        '        End Function
        '        '
        '        '=================================================================================
        '        '   Do not use this code
        '        '
        '        '   To retrieve a value from an option string, use csv.getAddonOption("name")
        '        '
        '        '   This was left here to work through any code issues that might arrise during
        '        '   the conversion.
        '        '
        '        '   Return the value from a name value pair, parsed with =,&[|].
        '        '   For example:
        '        '       name=Jay[Jay|Josh|Dwayne]
        '        '       the answer is Jay. If a select box is displayed, it is a dropdown of all three
        '        '=================================================================================
        '        '
        '        Public Function GetAggrOption_old(Name As String, SegmentCMDArgs As String) As String
        '            '
        '            Dim Pos as integer
        '            '
        '            GetAggrOption_old = getSimpleNameValue(Name, SegmentCMDArgs, "", vbCrLf)
        '            '
        '            ' remove the manual select list syntax "answer[choice1|choice2]"
        '            '
        '            Pos = InStr(1, GetAggrOption_old, "[")
        '            If Pos <> 0 Then
        '                GetAggrOption_old = Left(GetAggrOption_old, Pos - 1)
        '            End If
        '            '
        '            ' remove any function syntax "answer{selectcontentname RSS Feeds}"
        '            '
        '            Pos = InStr(1, GetAggrOption_old, "{")
        '            If Pos <> 0 Then
        '                GetAggrOption_old = Left(GetAggrOption_old, Pos - 1)
        '            End If
        '            '
        '        End Function
        '        '
        '        '=================================================================================
        '        '   Do not use this code
        '        '
        '        '   To retrieve a value from an option string, use csv.getAddonOption("name")
        '        '
        '        '   This was left here to work through any code issues that might arrise during
        '        '   Compatibility for GetArgument
        '        '=================================================================================
        '        '
        'Function kmaGetNameValue_old(Name As String, ArgumentString As String, Optional DefaultValue As String) As String
        '            kmaGetNameValue_old = getSimpleNameValue(Name, ArgumentString, DefaultValue, vbCrLf)
        '        End Function
        '        '
        '        '========================================================================
        '        '   KmaEncodeSQLText
        '        '========================================================================
        '        '
        '        Public Function KmaEncodeSQLText(ExpressionVariant As Object) As String
        '            'On Error GoTo ErrorTrap
        '            '
        '            'Dim MethodName As String
        '            '
        '            'MethodName = "KmaEncodeSQLText"
        '            '
        '            If IsNull(ExpressionVariant) Then
        '                KmaEncodeSQLText = "null"
        '            ElseIf IsMissing(ExpressionVariant) Then
        '                KmaEncodeSQLText = "null"
        '            ElseIf ExpressionVariant = "" Then
        '                KmaEncodeSQLText = "null"
        '            Else
        '                KmaEncodeSQLText = CStr(ExpressionVariant)
        '                ' ??? this should not be here -- to correct a field used in a CDef, truncate in SaveCS by fieldtype
        '                'KmaEncodeSQLText = Left(ExpressionVariant, 255)
        '                'remove-can not find a case where | is not allowed to be saved.
        '                'KmaEncodeSQLText = Replace(KmaEncodeSQLText, "|", "_")
        '                KmaEncodeSQLText = "'" & Replace(KmaEncodeSQLText, "'", "''") & "'"
        '            End If
        '            Exit Function
        '            '
        '            ' ----- Error Trap
        '            '
        'ErrorTrap:
        '        End Function
        '        '
        '        '========================================================================
        '        '   KmaEncodeSQLLongText
        '        '========================================================================
        '        '
        '        Public Function KmaEncodeSQLLongText(ExpressionVariant As Object) As String
        '            'On Error GoTo ErrorTrap
        '            '
        '            'Dim MethodName As String
        '            '
        '            'MethodName = "KmaEncodeSQLLongText"
        '            '
        '            If IsNull(ExpressionVariant) Then
        '                KmaEncodeSQLLongText = "null"
        '            ElseIf IsMissing(ExpressionVariant) Then
        '                KmaEncodeSQLLongText = "null"
        '            ElseIf ExpressionVariant = "" Then
        '                KmaEncodeSQLLongText = "null"
        '            Else
        '                KmaEncodeSQLLongText = ExpressionVariant
        '                'KmaEncodeSQLLongText = Replace(ExpressionVariant, "|", "_")
        '                KmaEncodeSQLLongText = "'" & Replace(KmaEncodeSQLLongText, "'", "''") & "'"
        '            End If
        '            Exit Function
        '            '
        '            ' ----- Error Trap
        '            '
        'ErrorTrap:
        '        End Function
        '        '
        '        '========================================================================
        '        '   KmaEncodeSQLDate
        '        '       encode a date variable to go in an sql expression
        '        '========================================================================
        '        '
        '        Public Function KmaEncodeSQLDate(ExpressionVariant As Object) As String
        '            'On Error GoTo ErrorTrap
        '            '
        '            Dim TimeVar As Date
        '            Dim TimeValuething As Single
        '            Dim TimeHours as integer
        '            Dim TimeMinutes as integer
        '            Dim TimeSeconds as integer
        '            'Dim MethodName As String
        '            ''
        '            'MethodName = "KmaEncodeSQLDate"
        '            '
        '            If IsNull(ExpressionVariant) Then
        '                KmaEncodeSQLDate = "null"
        '            ElseIf IsMissing(ExpressionVariant) Then
        '                KmaEncodeSQLDate = "null"
        '            ElseIf ExpressionVariant = "" Then
        '                KmaEncodeSQLDate = "null"
        '            ElseIf IsDate(ExpressionVariant) Then
        '                TimeVar = CDate(ExpressionVariant)
        '                If TimeVar = 0 Then
        '                    KmaEncodeSQLDate = "null"
        '                Else
        '                    TimeValuething = 86400.0! * (TimeVar - Int(TimeVar + 0.000011!))
        '                    TimeHours = Int(TimeValuething / 3600.0!)
        '                    If TimeHours >= 24 Then
        '                        TimeHours = 23
        '                    End If
        '                    TimeMinutes = Int(TimeValuething / 60.0!) - (TimeHours * 60)
        '                    If TimeMinutes >= 60 Then
        '                        TimeMinutes = 59
        '                    End If
        '                    TimeSeconds = TimeValuething - (TimeHours * 3600.0!) - (TimeMinutes * 60.0!)
        '                    If TimeSeconds >= 60 Then
        '                        TimeSeconds = 59
        '                    End If
        '                    KmaEncodeSQLDate = "{ts '" & Year(ExpressionVariant) & "-" & Right("0" & Month(ExpressionVariant), 2) & "-" & Right("0" & Day(ExpressionVariant), 2) & " " & Right("0" & TimeHours, 2) & ":" & Right("0" & TimeMinutes, 2) & ":" & Right("0" & TimeSeconds, 2) & "'}"
        '                End If
        '            Else
        '                KmaEncodeSQLDate = "null"
        '            End If
        '            Exit Function
        '            '
        '            ' ----- Error Trap
        '            '
        'ErrorTrap:
        '        End Function
        '        '
        '        '========================================================================
        '        '   KmaEncodeSQLNumber
        '        '       encode a number variable to go in an sql expression
        '        '========================================================================
        '        '
        '        Function KmaEncodeSQLNumber(ExpressionVariant As Object) As String
        '            'On Error GoTo ErrorTrap
        '            '
        '            'Dim MethodName As String
        '            ''
        '            'MethodName = "KmaEncodeSQLNumber"
        '            '
        '            If IsNull(ExpressionVariant) Then
        '                KmaEncodeSQLNumber = "null"
        '            ElseIf IsMissing(ExpressionVariant) Then
        '                KmaEncodeSQLNumber = "null"
        '            ElseIf ExpressionVariant = "" Then
        '                KmaEncodeSQLNumber = "null"
        '            ElseIf IsNumeric(ExpressionVariant) Then
        '                Select Case VarType(ExpressionVariant)
        '                    Case vbBoolean
        '                        If ExpressionVariant Then
        '                            KmaEncodeSQLNumber = SQLTrue
        '                        Else
        '                            KmaEncodeSQLNumber = SQLFalse
        '                        End If
        '                    Case Else
        '                        KmaEncodeSQLNumber = ExpressionVariant
        '                End Select
        '            Else
        '                KmaEncodeSQLNumber = "null"
        '            End If
        '            Exit Function
        '            '
        '            ' ----- Error Trap
        '            '
        'ErrorTrap:
        '        End Function
        '        '
        '        '========================================================================
        '        '   KmaEncodeSQLBoolean
        '        '       encode a boolean variable to go in an sql expression
        '        '========================================================================
        '        '
        '        Public Function KmaEncodeSQLBoolean(ExpressionVariant As Object) As String
        '            '
        '            KmaEncodeSQLBoolean = SQLFalse
        '            If Not IsNull(ExpressionVariant) Then
        '                If Not IsMissing(ExpressionVariant) Then
        '                    If ExpressionVariant <> False Then
        '                        KmaEncodeSQLBoolean = SQLTrue
        '                    End If
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        '   Gets the next line from a string, and removes the line
        '        '========================================================================
        '        '
        '        Public Function KmaGetLine(Body As String) As String
        '            Dim EOL As String
        '            Dim NextCR as integer
        '            Dim NextLF as integer
        '            Dim BOL as integer
        '            '
        '            NextCR = InStr(1, Body, vbCr)
        '            NextLF = InStr(1, Body, vbLf)

        '            If NextCR <> 0 Or NextLF <> 0 Then
        '                If NextCR <> 0 Then
        '                    If NextLF <> 0 Then
        '                        If NextCR < NextLF Then
        '                            EOL = NextCR - 1
        '                            If NextLF = NextCR + 1 Then
        '                                BOL = NextLF + 1
        '                            Else
        '                                BOL = NextCR + 1
        '                            End If

        '                        Else
        '                            EOL = NextLF - 1
        '                            BOL = NextLF + 1
        '                        End If
        '                    Else
        '                        EOL = NextCR - 1
        '                        BOL = NextCR + 1
        '                    End If
        '                Else
        '                    EOL = NextLF - 1
        '                    BOL = NextLF + 1
        '                End If
        '                KmaGetLine = Mid(Body, 1, EOL)
        '                Body = Mid(Body, BOL)
        '            Else
        '                KmaGetLine = Body
        '                Body = ""
        '            End If

        '            'EOL = InStr(1, Body, vbCrLf)

        '            'If EOL <> 0 Then
        '            '    KmaGetLine = Mid(Body, 1, EOL - 1)
        '            '    Body = Mid(Body, EOL + 2)
        '            '    End If
        '            '
        '        End Function
        '        '
        '        '=================================================================================
        '        '   Get a Random Long Value
        '        '=================================================================================
        '        '
        '        Public Function GetRandomInteger() as integer
        '            '
        '            Dim RandomBase as integer
        '            Dim RandomLimit as integer
        '            '
        '            RandomBase = App.ThreadID
        '            RandomBase = RandomBase And ((2 ^ 30) - 1)
        '            RandomLimit = (2 ^ 31) - RandomBase - 1
        '            Randomize()
        '            GetRandomInteger = RandomBase + (Rnd() * RandomLimit)
        '            '
        '        End Function
        '        '
        '        '=================================================================================
        '        '
        '        '=================================================================================
        '        '
        '        Public Function IsRSOK(RS As Object) As Boolean
        '            IsRSOK = False
        '            If Not (RS Is Nothing) Then
        '                If RS.State <> 0 Then
        '                    If Not RS.EOF Then
        '                        IsRSOK = True
        '                    End If
        '                End If
        '            End If
        '        End Function
        '        '
        '        '=================================================================================
        '        '
        '        '=================================================================================
        '        '
        '        Public Sub CloseRS(RS As Object)
        '            If Not (RS Is Nothing) Then
        '                If RS.State <> 0 Then
        '                    Call RS.Close()
        '                End If
        '            End If
        '        End Sub
        '        '
        '        '=============================================================================
        '        ' Create the part of the sql where clause that is modified by the user
        '        '   WorkingQuery is the original querystring to change
        '        '   QueryName is the name part of the name pair to change
        '        '   If the QueryName is not found in the string, ignore call
        '        '=============================================================================
        '        '
        'Public Function ModifyQueryString(WorkingQuery As String, QueryName As String, QueryValue As String, Optional AddIfMissing As Variant) As String
        '            '
        '            If InStr(1, WorkingQuery, "?") Then
        '                ModifyQueryString = cp.Utils.ModifyLinkQueryString(WorkingQuery, QueryName, QueryValue, AddIfMissing)
        '            Else
        '                ModifyQueryString = Mid(cp.Utils.ModifyLinkQueryString("?" & WorkingQuery, QueryName, QueryValue, AddIfMissing), 2)
        '            End If
        '        End Function
        '        '
        '        '=============================================================================
        '        '   Modify a querystring name/value pair in a Link
        '        '=============================================================================
        '        '
        'Public Function cp.Utils.ModifyLinkQueryString(Link As String, QueryName As String, QueryValue As String, Optional AddIfMissing As Variant) As String
        '            '
        '            Dim PositionName as integer
        '            Dim PositionEqual as integer
        '            Dim PositionValueStart as integer
        '            Dim PositionValueEnd as integer
        '            Dim Element() As String
        '            Dim ElementCount as integer
        '            Dim ElementPointer as integer
        '            Dim NameValue() As String
        '            Dim UcaseQueryName As String
        '            Dim ElementFound As Boolean
        '            Dim iAddIfMissing As Boolean
        '            Dim LeftPart As String
        '            Dim QueryString As String
        '            '
        '            iAddIfMissing = KmaEncodeMissingBoolean(AddIfMissing, True)
        '            If InStr(1, Link, "?") <> 0 Then
        '                cp.Utils.ModifyLinkQueryString = Mid(Link, 1, InStr(1, Link, "?") - 1)
        '                QueryString = Mid(Link, Len(cp.Utils.ModifyLinkQueryString) + 2)
        '            Else
        '                cp.Utils.ModifyLinkQueryString = Link
        '                QueryString = ""
        '            End If
        '            UcaseQueryName = UCase(kmaEncodeRequestVariable(QueryName))
        '            If QueryString <> "" Then
        '                Element = Split(QueryString, "&")
        '                ElementCount = UBound(Element) + 1
        '                For ElementPointer = 0 To ElementCount - 1
        '                    NameValue = Split(Element(ElementPointer), "=")
        '                    If UBound(NameValue) = 1 Then
        '                        If UCase(NameValue(0)) = UcaseQueryName Then
        '                            If QueryValue = "" Then
        '                                Element(ElementPointer) = ""
        '                            Else
        '                                Element(ElementPointer) = QueryName & "=" & QueryValue
        '                            End If
        '                            ElementFound = True
        '                            Exit For
        '                        End If
        '                    End If
        '                Next
        '            End If
        '            If Not ElementFound And (QueryValue <> "") Then
        '                '
        '                ' element not found, it needs to be added
        '                '
        '                If iAddIfMissing Then
        '                    If QueryString = "" Then
        '                        QueryString =cp.Utils.EncodeRequestVariable(QueryName) & "=" &cp.Utils.EncodeRequestVariable(QueryValue)
        '                    Else
        '                        QueryString = QueryString & "&" &cp.Utils.EncodeRequestVariable(QueryName) & "=" &cp.Utils.EncodeRequestVariable(QueryValue)
        '                    End If
        '                End If
        '            Else
        '                '
        '                ' element found
        '                '
        '                QueryString = Join(Element, "&")
        '                If (QueryString <> "") And (QueryValue = "") Then
        '                    '
        '                    ' element found and needs to be removed
        '                    '
        '                    QueryString = Replace(QueryString, "&&", "&")
        '                    If Left(QueryString, 1) = "&" Then
        '                        QueryString = Mid(QueryString, 2)
        '                    End If
        '                    If Right(QueryString, 1) = "&" Then
        '                        QueryString = Mid(QueryString, 1, Len(QueryString) - 1)
        '                    End If
        '                End If
        '            End If
        '            If (QueryString <> "") Then
        '                cp.Utils.ModifyLinkQueryString = cp.Utils.ModifyLinkQueryString & "?" & QueryString
        '            End If
        '        End Function
        '        '
        '        '=================================================================================
        '        '
        '        '=================================================================================
        '        '
        '        Public Function GetIntegerString(Value as integer, DigitCount as integer) As String
        '            If Len(Value) <= DigitCount Then
        '        GetIntegerString = String(DigitCount - Len(CStr(Value)), "0") & CStr(Value)
        '            Else
        '                GetIntegerString = CStr(Value)
        '            End If
        '        End Function
        '        '
        '        '==========================================================================================
        '        '   Set the current process to a high priority
        '        '       Should be called once from the objects parent when it is first created.
        '        '
        '        '   taken from an example labeled
        '        '       KPD-Team 2000
        '        '       URL: http://www.allapi.net/
        '        '       Email: KPDTeam@Allapi.net
        '        '==========================================================================================
        '        '
        '        Public Sub SetProcessHighPriority()
        '            Dim hProcess as integer
        '            '
        '            'set the new priority class
        '            '
        '            hProcess = GetCurrentProcess
        '            Call SetPriorityClass(hProcess, HIGH_PRIORITY_CLASS)
        '            '
        '        End Sub
        '        '
        '        '==========================================================================================
        '        '   Format the current error object into a standard string
        '        '==========================================================================================
        '        '
        'Public Function GetErrString(Optional ErrorObject As Object) As String
        '            Dim Copy As String
        '            If ErrorObject Is Nothing Then
        '                If Err.Number = 0 Then
        '                    GetErrString = "[no error]"
        '                Else
        '                    Copy = Err.Description
        '                    Copy = Replace(Copy, vbCrLf, "-")
        '                    Copy = Replace(Copy, vbLf, "-")
        '                    Copy = Replace(Copy, vbCrLf, "")
        '                    GetErrString = "[" & Err.Source & " #" & Err.Number & ", " & Copy & "]"
        '                End If
        '            Else
        '                If ErrorObject.Number = 0 Then
        '                    GetErrString = "[no error]"
        '                Else
        '                    Copy = ErrorObject.Description
        '                    Copy = Replace(Copy, vbCrLf, "-")
        '                    Copy = Replace(Copy, vbLf, "-")
        '                    Copy = Replace(Copy, vbCrLf, "")
        '                    GetErrString = "[" & ErrorObject.Source & " #" & ErrorObject.Number & ", " & Copy & "]"
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        '==========================================================================================
        '        '   Format the current error object into a standard string
        '        '==========================================================================================
        '        '
        '        Public Function GetProcessID() as integer
        '            GetProcessID = GetCurrentProcessId
        '        End Function
        '        '
        '        '==========================================================================================
        '        '   Test if a test string is in a delimited string
        '        '==========================================================================================
        '        '
        '        Public Function IsInDelimitedString(DelimitedString As String, TestString As String, Delimiter As String) As Boolean
        '            IsInDelimitedString = (0 <> InStr(1, Delimiter & DelimitedString & Delimiter, Delimiter & TestString & Delimiter, vbTextCompare))
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaEncodeURL
        '        '
        '        '   Encodes only what is to the left of the first ?
        '        '   All URL path characters are assumed to be correct (/:#)
        '        '========================================================================
        '        '
        '        Function kmaEncodeURL(Source As String) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim URLSplit() As String
        '            Dim LeftSide As String
        '            Dim RightSide As String
        '            '
        '            kmaEncodeURL = Source
        '            If Source <> "" Then
        '                URLSplit = Split(Source, "?")
        '                kmaEncodeURL = URLSplit(0)
        '                kmaEncodeURL = Replace(kmaEncodeURL, "%", "%25")
        '                '
        '                kmaEncodeURL = Replace(kmaEncodeURL, """", "%22")
        '                kmaEncodeURL = Replace(kmaEncodeURL, " ", "%20")
        '                kmaEncodeURL = Replace(kmaEncodeURL, "$", "%24")
        '                kmaEncodeURL = Replace(kmaEncodeURL, "+", "%2B")
        '                kmaEncodeURL = Replace(kmaEncodeURL, ",", "%2C")
        '                kmaEncodeURL = Replace(kmaEncodeURL, ";", "%3B")
        '                kmaEncodeURL = Replace(kmaEncodeURL, "<", "%3C")
        '                kmaEncodeURL = Replace(kmaEncodeURL, "=", "%3D")
        '                kmaEncodeURL = Replace(kmaEncodeURL, ">", "%3E")
        '                kmaEncodeURL = Replace(kmaEncodeURL, "@", "%40")
        '                If UBound(URLSplit) > 0 Then
        '                    kmaEncodeURL = kmaEncodeURL & "?" & kmaEncodeQueryString(URLSplit(1))
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaEncodeQueryString
        '        '
        '        '   This routine encodes the URL QueryString to conform to rules
        '        '========================================================================
        '        '
        '        Function kmaEncodeQueryString(Source As String) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim QSSplit() As String
        '            Dim QSPointer as integer
        '            Dim NVSplit() As String
        '            Dim NV As String
        '            '
        '            kmaEncodeQueryString = ""
        '            If Source <> "" Then
        '                QSSplit = Split(Source, "&")
        '                For QSPointer = 0 To UBound(QSSplit)
        '                    NV = QSSplit(QSPointer)
        '                    If NV <> "" Then
        '                        NVSplit = Split(NV, "=")
        '                        If UBound(NVSplit) = 0 Then
        '                            NVSplit(0) =cp.Utils.EncodeRequestVariable(NVSplit(0))
        '                            kmaEncodeQueryString = kmaEncodeQueryString & "&" & NVSplit(0)
        '                        Else
        '                            NVSplit(0) =cp.Utils.EncodeRequestVariable(NVSplit(0))
        '                            NVSplit(1) =cp.Utils.EncodeRequestVariable(NVSplit(1))
        '                            kmaEncodeQueryString = kmaEncodeQueryString & "&" & NVSplit(0) & "=" & NVSplit(1)
        '                        End If
        '                    End If
        '                Next
        '                If kmaEncodeQueryString <> "" Then
        '                    kmaEncodeQueryString = Mid(kmaEncodeQueryString, 2)
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        'cp.Utils.EncodeRequestVariable
        '        '
        '        '   This routine encodes a request variable for a URL Query String
        '        '       ...can be the requestname or the requestvalue
        '        '========================================================================
        '        '
        '        Functioncp.Utils.EncodeRequestVariable(Source As String) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim SourcePointer as integer
        '            Dim Character As String
        '            Dim LocalSource As String
        '            '
        '            If Source <> "" Then
        '                LocalSource = Source
        '                ' "+" is an allowed character for filenames. If you add it, the wrong file will be looked up
        '                'LocalSource = Replace(LocalSource, " ", "+")
        '                For SourcePointer = 1 To Len(LocalSource)
        '                    Character = Mid(LocalSource, SourcePointer, 1)
        '                    ' "%" added so if this is called twice, it will not destroy "%20" values
        '                    'If Character = " " Then
        '                    '   cp.Utils.EncodeRequestVariable =cp.Utils.EncodeRequestVariable & "+"
        '                    If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_!*()", Character, vbTextCompare) <> 0 Then
        '                        'ElseIf InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789./:-_!*()", Character, vbTextCompare) <> 0 Then
        '                        'ElseIf InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789./:?#-_!~*'()%", Character, vbTextCompare) <> 0 Then
        '                       cp.Utils.EncodeRequestVariable =cp.Utils.EncodeRequestVariable & Character
        '                    Else
        '                       cp.Utils.EncodeRequestVariable =cp.Utils.EncodeRequestVariable & "%" & Hex(Asc(Character))
        '                    End If
        '                Next
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' cp.Utils.EncodeHTML
        '        '
        '        '   Convert all characters that are not allowed in HTML to their Text equivalent
        '        '   in preperation for use on an HTML page
        '        '========================================================================
        '        '
        '        Function cp.Utils.EncodeHTML(Source As String) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            cp.Utils.EncodeHTML = Source
        '            cp.Utils.EncodeHTML = Replace(cp.Utils.EncodeHTML, "&", "&amp;")
        '            cp.Utils.EncodeHTML = Replace(cp.Utils.EncodeHTML, "<", "&lt;")
        '            cp.Utils.EncodeHTML = Replace(cp.Utils.EncodeHTML, ">", "&gt;")
        '            cp.Utils.EncodeHTML = Replace(cp.Utils.EncodeHTML, """", "&quot;")
        '            cp.Utils.EncodeHTML = Replace(cp.Utils.EncodeHTML, "'", "&#39;")
        '            'cp.Utils.EncodeHTML = Replace(cp.Utils.EncodeHTML, "'", "&apos;")
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaDecodeHTML
        '        '
        '        '   Convert HTML equivalent characters to their equivalents
        '        '========================================================================
        '        '
        '        Function kmaDecodeHTML(Source As String) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim Pos as integer
        '            Dim s As String
        '            Dim CharCodeString As String
        '            Dim CharCode as integer
        '            Dim posEnd as integer
        '            '
        '            ' 11/26/2009 - basically re-wrote it, I commented the old one out below
        '            '
        '            kmaDecodeHTML = ""
        '            If Source <> "" Then
        '                s = Source
        '                '
        '                Pos = Len(s)
        '                Pos = InStrRev(s, "&#", Pos)
        '                Do While Pos <> 0
        '                    CharCodeString = ""
        '                    If Mid(s, Pos + 3, 1) = ";" Then
        '                        CharCodeString = Mid(s, Pos + 2, 1)
        '                        posEnd = Pos + 4
        '                    ElseIf Mid(s, Pos + 4, 1) = ";" Then
        '                        CharCodeString = Mid(s, Pos + 2, 2)
        '                        posEnd = Pos + 5
        '                    ElseIf Mid(s, Pos + 5, 1) = ";" Then
        '                        CharCodeString = Mid(s, Pos + 2, 3)
        '                        posEnd = Pos + 6
        '                    End If
        '                    If CharCodeString <> "" Then
        '                        If IsNumeric(CharCodeString) Then
        '                            CharCode = CLng(CharCodeString)
        '                            s = Mid(s, 1, Pos - 1) & Chr(CharCode) & Mid(s, posEnd)
        '                        End If
        '                    End If
        '                    '
        '                    Pos = InStrRev(s, "&#", Pos)
        '                Loop
        '                '
        '                ' replace out all common names (at least the most common for now)
        '                '
        '                s = Replace(s, "&lt;", "<")
        '                s = Replace(s, "&gt;", ">")
        '                s = Replace(s, "&quot;", """")
        '                s = Replace(s, "&apos;", "'")
        '                '
        '                ' Always replace the amp last
        '                '
        '                s = Replace(s, "&amp;", "&")
        '                '
        '                kmaDecodeHTML = s
        '            End If
        '            ' pre-11/26/2009
        '            'kmaDecodeHTML = Source
        '            'kmaDecodeHTML = Replace(kmaDecodeHTML, "&amp;", "&")
        '            'kmaDecodeHTML = Replace(kmaDecodeHTML, "&lt;", "<")
        '            'kmaDecodeHTML = Replace(kmaDecodeHTML, "&gt;", ">")
        '            'kmaDecodeHTML = Replace(kmaDecodeHTML, "&quot;", """")
        '            'kmaDecodeHTML = Replace(kmaDecodeHTML, "&nbsp;", " ")
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaAddSpanClass
        '        '
        '        '   Adds a span around the copy with the class name provided
        '        '========================================================================
        '        '
        '        Function kmaAddSpan(Copy As String, ClassName As String) As String
        '            '
        '            kmaAddSpan = "<SPAN Class=""" & ClassName & """>" & Copy & "</SPAN>"
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaDecodeResponseVariable
        '        '
        '        '   Converts a querystring name or value back into the characters it represents
        '        '   This is the same code as the decodeurl
        '        '========================================================================
        '        '
        '        Function kmaDecodeResponseVariable(Source As String) As String
        '            '
        '            Dim Position as integer
        '            Dim ESCString As String
        '            Dim ESCValue as integer
        '            Dim Digit0 As String
        '            Dim Digit1 As String
        '            Dim iURL As String
        '            '
        '            iURL = kmaEncodeText(Source)
        '            kmaDecodeResponseVariable = Replace(iURL, "+", " ")
        '            Position = InStr(1, kmaDecodeResponseVariable, "%")
        '            Do While Position <> 0
        '                ESCString = Mid(kmaDecodeResponseVariable, Position, 3)
        '                Digit0 = UCase(Mid(ESCString, 2, 1))
        '                Digit1 = UCase(Mid(ESCString, 3, 1))
        '                If ((Digit0 >= "0") And (Digit0 <= "9")) Or ((Digit0 >= "A") And (Digit0 <= "F")) Then
        '                    If ((Digit1 >= "0") And (Digit1 <= "9")) Or ((Digit1 >= "A") And (Digit1 <= "F")) Then
        '                        ESCValue = CLng("&H" & Mid(ESCString, 2))
        '                        kmaDecodeResponseVariable = Mid(kmaDecodeResponseVariable, 1, Position - 1) & Chr(ESCValue) & Mid(kmaDecodeResponseVariable, Position + 3)
        '                        '  & Replace(kmaDecodeResponseVariable, ESCString, Chr(ESCValue), Position, 1)
        '                    End If
        '                End If
        '                Position = InStr(Position + 1, kmaDecodeResponseVariable, "%")
        '            Loop
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmadecodeURL
        '        '   Converts a querystring from an Encoded URL (with %20 and +), to non incoded (with spaced)
        '        '========================================================================
        '        '
        '        Function kmaDecodeURL(Source As String) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim Position as integer
        '            Dim ESCString As String
        '            Dim ESCValue as integer
        '            Dim Digit0 As String
        '            Dim Digit1 As String
        '            Dim iURL As String
        '            '
        '            iURL = kmaEncodeText(Source)
        '            kmaDecodeURL = Replace(iURL, "+", " ")
        '            Position = InStr(1, kmaDecodeURL, "%")
        '            Do While Position <> 0
        '                ESCString = Mid(kmaDecodeURL, Position, 3)
        '                Digit0 = UCase(Mid(ESCString, 2, 1))
        '                Digit1 = UCase(Mid(ESCString, 3, 1))
        '                If ((Digit0 >= "0") And (Digit0 <= "9")) Or ((Digit0 >= "A") And (Digit0 <= "F")) Then
        '                    If ((Digit1 >= "0") And (Digit1 <= "9")) Or ((Digit1 >= "A") And (Digit1 <= "F")) Then
        '                        ESCValue = CLng("&H" & Mid(ESCString, 2))
        '                        kmaDecodeURL = Replace(kmaDecodeURL, ESCString, Chr(ESCValue))
        '                    End If
        '                End If
        '                Position = InStr(Position + 1, kmaDecodeURL, "%")
        '            Loop
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaGetFirstNonZeroDate
        '        '
        '        '   Converts a querystring name or value back into the characters it represents
        '        '========================================================================
        '        '
        '        Function kmaGetFirstNonZeroDate(Date0 As Date, Date1 As Date) As Date
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim NullDate As Date
        '            '
        '            NullDate = CDate(0)
        '            If Date0 = NullDate Then
        '                If Date1 = NullDate Then
        '                    '
        '                    ' Both 0, return 0
        '                    '
        '                    kmaGetFirstNonZeroDate = NullDate
        '                Else
        '                    '
        '                    ' Date0 is NullDate, return Date1
        '                    '
        '                    kmaGetFirstNonZeroDate = Date1
        '                End If
        '            Else
        '                If Date1 = NullDate Then
        '                    '
        '                    ' Date1 is nulldate, return Date0
        '                    '
        '                    kmaGetFirstNonZeroDate = Date0
        '                ElseIf Date0 < Date1 Then
        '                    '
        '                    ' Date0 is first
        '                    '
        '                    kmaGetFirstNonZeroDate = Date0
        '                Else
        '                    '
        '                    ' Date1 is first
        '                    '
        '                    kmaGetFirstNonZeroDate = Date1
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaGetFirstposition
        '        '
        '        '   returns 0 if both are zero
        '        '   returns 1 if the first integer is non-zero and less then the second
        '        '   returns 2 if the second integer is non-zero and less then the first
        '        '========================================================================
        '        '
        '        Function kmaGetFirstNonZeroLong(Integer1 as integer, Integer2 as integer) as integer
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            If Integer1 = 0 Then
        '                If Integer2 = 0 Then
        '                    '
        '                    ' Both 0, return 0
        '                    '
        '                    kmaGetFirstNonZeroLong = 0
        '                Else
        '                    '
        '                    ' Integer1 is 0, return Integer2
        '                    '
        '                    kmaGetFirstNonZeroLong = 2
        '                End If
        '            Else
        '                If Integer2 = 0 Then
        '                    '
        '                    ' Integer2 is 0, return Integer1
        '                    '
        '                    kmaGetFirstNonZeroLong = 1
        '                ElseIf Integer1 < Integer2 Then
        '                    '
        '                    ' Integer1 is first
        '                    '
        '                    kmaGetFirstNonZeroLong = 1
        '                Else
        '                    '
        '                    ' Integer2 is first
        '                    '
        '                    kmaGetFirstNonZeroLong = 2
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaSplit
        '        '   returns the result of a Split, except it honors quoted text
        '        '   if a quote is found, it is assumed to also be a delimiter ( 'this"that"theother' = 'this "that" theother' )
        '        '========================================================================
        '        '
        '        Function kmaSplit(WordList As String, Delimiter As String) As Object
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim QuoteSplit() As String
        '            Dim QuoteSplitCount as integer
        '            Dim QuoteSplitPointer as integer
        '            Dim InQuote As Boolean
        '            Dim Out() As String
        '            Dim OutPointer as integer
        '            Dim OutSize as integer
        '            Dim SpaceSplit() As String
        '            Dim SpaceSplitCount as integer
        '            Dim SpaceSplitPointer as integer
        '            Dim Fragment As String
        '            '
        '            OutPointer = 0
        '            ReDim Out(0)
        '            OutSize = 1
        '            If WordList <> "" Then
        '                QuoteSplit = Split(WordList, """")
        '                QuoteSplitCount = UBound(QuoteSplit) + 1
        '                InQuote = (Mid(WordList, 1, 1) = "")
        '                For QuoteSplitPointer = 0 To QuoteSplitCount - 1
        '                    Fragment = QuoteSplit(QuoteSplitPointer)
        '                    If Fragment = "" Then
        '                        '
        '                        ' empty fragment
        '                        ' this is a quote at the end, or two quotes together
        '                        ' do not skip to the next out pointer
        '                        '
        '                        If OutPointer >= OutSize Then
        '                            OutSize = OutSize + 10
        '                            ReDim Preserve Out(OutSize)
        '                        End If
        '                        'OutPointer = OutPointer + 1
        '                    Else
        '                        If Not InQuote Then
        '                            SpaceSplit = Split(Fragment, Delimiter)
        '                            SpaceSplitCount = UBound(SpaceSplit) + 1
        '                            For SpaceSplitPointer = 0 To SpaceSplitCount - 1
        '                                If OutPointer >= OutSize Then
        '                                    OutSize = OutSize + 10
        '                                    ReDim Preserve Out(OutSize)
        '                                End If
        '                                Out(OutPointer) = Out(OutPointer) & SpaceSplit(SpaceSplitPointer)
        '                                If (SpaceSplitPointer <> (SpaceSplitCount - 1)) Then
        '                                    '
        '                                    ' divide output between splits
        '                                    '
        '                                    OutPointer = OutPointer + 1
        '                                    If OutPointer >= OutSize Then
        '                                        OutSize = OutSize + 10
        '                                        ReDim Preserve Out(OutSize)
        '                                    End If
        '                                End If
        '                            Next
        '                        Else
        '                            Out(OutPointer) = Out(OutPointer) & """" & Fragment & """"
        '                        End If
        '                    End If
        '                    InQuote = Not InQuote
        '                Next
        '            End If
        '            ReDim Preserve Out(OutPointer)
        '            '
        '            '
        '            kmaSplit = Out
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' kmaSplit_Old
        '        '   returns the result of a Split, except it honors quoted text
        '        '   if a quote is found, it is assumed to also be a delimiter ( 'this"that"theother' = 'this "that" theother' )
        '        '========================================================================
        '        '
        '        Function kmaSplit_Old(WordList As String, Delimiter As String) As Object
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim CurrentPosition as integer
        '            Dim NextDelimiterPosition as integer
        '            Dim NextQuotePosition as integer
        '            Dim ResultCount as integer
        '            Dim Result() As String
        '            Dim WorkingWordList As String
        '            Dim LenWorkingWordList as integer
        '            '
        '            WorkingWordList = Trim(WordList)
        '            If WorkingWordList <> "" Then
        '                LenWorkingWordList = Len(WorkingWordList)
        '                CurrentPosition = 1
        '                ResultCount = 0
        '                Do While CurrentPosition <> 0
        '                    NextDelimiterPosition = InStr(CurrentPosition, WorkingWordList, Delimiter, vbTextCompare)
        '                    NextQuotePosition = InStr(CurrentPosition, WorkingWordList, """", vbTextCompare)
        '                    Select Case kmaGetFirstNonZeroLong(NextDelimiterPosition, NextQuotePosition)
        '                        Case 0
        '                            '
        '                            ' no more left
        '                            '
        '                            If CurrentPosition <= LenWorkingWordList Then
        '                                ReDim Preserve Result(ResultCount)
        '                                Result(ResultCount) = Mid(WorkingWordList, CurrentPosition)
        '                                ResultCount = ResultCount + 1
        '                            End If
        '                            CurrentPosition = 0
        '                        Case 1
        '                            '
        '                            ' Delimiter found before quote
        '                            '
        '                            ReDim Preserve Result(ResultCount)
        '                            Result(ResultCount) = Mid(WorkingWordList, CurrentPosition, NextDelimiterPosition - CurrentPosition)
        '                            ResultCount = ResultCount + 1
        '                            If NextDelimiterPosition >= LenWorkingWordList Then
        '                                CurrentPosition = 0
        '                            Else
        '                                CurrentPosition = NextDelimiterPosition + 1
        '                            End If
        '                        Case 2
        '                            '
        '                            ' Quote Found before delimiter
        '                            '
        '                            CurrentPosition = NextQuotePosition + 1
        '                            NextQuotePosition = InStr(CurrentPosition, WorkingWordList, """", vbTextCompare)
        '                            If NextQuotePosition = 0 Then
        '                                '
        '                                ' Problem, as single quote. Just end the phrase here
        '                                '
        '                                NextQuotePosition = LenWorkingWordList + 1
        '                            End If
        '                            ReDim Preserve Result(ResultCount)
        '                            Result(ResultCount) = Mid(WorkingWordList, CurrentPosition, NextQuotePosition - CurrentPosition)
        '                            ResultCount = ResultCount + 1
        '                            If NextQuotePosition >= LenWorkingWordList Then
        '                                CurrentPosition = 0
        '                            Else
        '                                CurrentPosition = NextQuotePosition + 1
        '                            End If
        '                    End Select
        '                    '
        '                    ' pass any delimiters
        '                    '
        '                    If CurrentPosition <> 0 Then
        '                        Do While Mid(WorkingWordList, CurrentPosition, 1) = Delimiter
        '                            CurrentPosition = CurrentPosition + 1
        '                            If CurrentPosition >= LenWorkingWordList Then
        '                                CurrentPosition = 0
        '                                Exit Do
        '                            End If
        '                        Loop
        '                    End If
        '                Loop
        '            End If
        '            kmaSplit_Old = Result
        '            '

        '            '
        '            'kmaSplit_Old = Split(WorkingWordList, Delimiter, , vbTextCompare)
        '            '
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaGetYesNo(Key As Boolean) As String
        '            If Key Then
        '                kmaGetYesNo = "Yes"
        '            Else
        '                kmaGetYesNo = "No"
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaGetFilename(PathFilename As String) As String
        '            Dim Position as integer
        '            '
        '            kmaGetFilename = PathFilename
        '            Position = InStrRev(kmaGetFilename, "/")
        '            If Position <> 0 Then
        '                kmaGetFilename = Mid(kmaGetFilename, Position + 1)
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        'Public Function kmaStartTable(Padding as integer, Spacing as integer, Border as integer, Optional ClassStyle As String) As String
        '            kmaStartTable = "<table border=""" & Border & """ cellpadding=""" & Padding & """ cellspacing=""" & Spacing & """ class=""" & ClassStyle & """ width=""100%"">"
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaStartTableRow() As String
        '            kmaStartTableRow = "<tr>"
        '        End Function
        '        '
        '        '
        '        '
        'Public Function kmaStartTableCell(Optional Width As String, Optional ColSpan as integer, Optional EvenRow As Boolean, Optional Align As String, Optional BGColor As String) As String
        '            If Width <> "" Then
        '                kmaStartTableCell = " width=""" & Width & """"
        '            End If
        '            If BGColor <> "" Then
        '                kmaStartTableCell = kmaStartTableCell & " bgcolor=""" & BGColor & """"
        '            ElseIf EvenRow Then
        '                kmaStartTableCell = kmaStartTableCell & " class=""ccPanelRowEven"""
        '            Else
        '                kmaStartTableCell = kmaStartTableCell & " class=""ccPanelRowOdd"""
        '            End If
        '            If ColSpan <> 0 Then
        '                kmaStartTableCell = kmaStartTableCell & " colspan=""" & ColSpan & """"
        '            End If
        '            If Align <> "" Then
        '                kmaStartTableCell = kmaStartTableCell & " align=""" & Align & """"
        '            End If
        '            kmaStartTableCell = "<TD" & kmaStartTableCell & ">"
        '        End Function
        '        '
        '        '
        '        '
        'Public Function kmaGetTableCell(Copy As String, Optional Width As String, Optional ColSpan as integer, Optional EvenRow As Boolean, Optional Align As String, Optional BGColor As String) As String
        '            kmaGetTableCell = kmaStartTableCell(Width, ColSpan, EvenRow, Align, BGColor) & Copy & kmaEndTableCell
        '        End Function
        '        '
        '        '
        '        '
        'Public Function kmaGetTableRow(Cell As String, Optional ColSpan as integer, Optional EvenRow As Boolean) As String
        '            kmaGetTableRow = kmaStartTableRow() & kmaGetTableCell(Cell, "100%", ColSpan, EvenRow) & kmaEndTableRow
        '        End Function
        '        '
        '        ' remove the host and approotpath, leaving the "active" path and all else
        '        '
        '        Public Function kmaConvertShortLinkToLink(URL As String, PathPagePrefix As String) As String
        '            kmaConvertShortLinkToLink = URL
        '            If URL <> "" And PathPagePrefix <> "" Then
        '                If InStr(1, kmaConvertShortLinkToLink, PathPagePrefix, vbTextCompare) = 1 Then
        '                    kmaConvertShortLinkToLink = Mid(kmaConvertShortLinkToLink, Len(PathPagePrefix) + 1)
        '                End If
        '            End If
        '        End Function
        '        '
        '        ' ------------------------------------------------------------------------------------------------------
        '        '   Preserve URLs that do not start HTTP or HTTPS
        '        '   Preserve URLs from other sites (offsite)
        '        '   Preserve HTTP://ServerHost/ServerVirtualPath/Files/ in all cases
        '        '   Convert HTTP://ServerHost/ServerVirtualPath/folder/page -> /folder/page
        '        '   Convert HTTP://ServerHost/folder/page -> /folder/page
        '        ' ------------------------------------------------------------------------------------------------------
        '        '
        '        Public Function kmaConvertLinkToShortLink(URL As String, ServerHost As String, ServerVirtualPath As String) As String
        '            '
        '            Dim BadString As String
        '            Dim GoodString As String
        '            Dim Protocol As String
        '            Dim WorkingLink As String
        '            '
        '            WorkingLink = URL
        '            '
        '            ' ----- Determine Protocol
        '            '
        '            If InStr(1, WorkingLink, "HTTP://", vbTextCompare) = 1 Then
        '                '
        '                ' HTTP
        '                '
        '                Protocol = Mid(WorkingLink, 1, 7)
        '            ElseIf InStr(1, WorkingLink, "HTTPS://", vbTextCompare) = 1 Then
        '                '
        '                ' HTTPS
        '                '
        '                ' try this -- a ssl link can not be shortened
        '                kmaConvertLinkToShortLink = WorkingLink
        '                Exit Function
        '                Protocol = Mid(WorkingLink, 1, 8)
        '            End If
        '            If Protocol <> "" Then
        '                '
        '                ' ----- Protcol found, determine if is local
        '                '
        '                GoodString = Protocol & ServerHost
        '                If (InStr(1, WorkingLink, GoodString, vbTextCompare) <> 0) Then
        '                    '
        '                    ' URL starts with Protocol ServerHost
        '                    '
        '                    GoodString = Protocol & ServerHost & ServerVirtualPath & "/files/"
        '                    If (InStr(1, WorkingLink, GoodString, vbTextCompare) <> 0) Then
        '                        '
        '                        ' URL is in the virtual files directory
        '                        '
        '                        BadString = GoodString
        '                        GoodString = ServerVirtualPath & "/files/"
        '                        WorkingLink = Replace(WorkingLink, BadString, GoodString, , , vbTextCompare)
        '                    Else
        '                        '
        '                        ' URL is not in files virtual directory
        '                        '
        '                        BadString = Protocol & ServerHost & ServerVirtualPath & "/"
        '                        GoodString = "/"
        '                        WorkingLink = Replace(WorkingLink, BadString, GoodString, , , vbTextCompare)
        '                        '
        '                        BadString = Protocol & ServerHost & "/"
        '                        GoodString = "/"
        '                        WorkingLink = Replace(WorkingLink, BadString, GoodString, , , vbTextCompare)
        '                    End If
        '                End If
        '            End If
        '            kmaConvertLinkToShortLink = WorkingLink
        '        End Function
        '        '
        '        ' Correct the link for the virtual path, either add it or remove it
        '        '
        '        Public Function kmaEncodeAppRootPath(Link As String, VirtualPath As String, AppRootPath As String, ServerHost As String) As String
        '            '
        '            Dim Protocol As String
        '            Dim Host As String
        '            Dim Path As String
        '            Dim Page As String
        '            Dim QueryString As String
        '            Dim VirtualHosted As Boolean
        '            '
        '            kmaEncodeAppRootPath = Link
        '            If (InStr(1, kmaEncodeAppRootPath, ServerHost, vbTextCompare) <> 0) Or (InStr(1, Link, "/") = 1) Then
        '                'If (InStr(1, kmaEncodeAppRootPath, ServerHost, vbTextCompare) <> 0) And (InStr(1, Link, "/") <> 0) Then
        '                '
        '                ' This link is onsite and has a path
        '                '
        '                VirtualHosted = (InStr(1, AppRootPath, VirtualPath, vbTextCompare) <> 0)
        '                If VirtualHosted And (InStr(1, Link, AppRootPath, vbTextCompare) = 1) Then
        '                    '
        '                    ' quick - virtual hosted and link starts at AppRootPath
        '                    '
        '                ElseIf (Not VirtualHosted) And (Mid(Link, 1, 1) = "/") And (InStr(1, Link, AppRootPath, vbTextCompare) = 1) Then
        '                    '
        '                    ' quick - not virtual hosted and link starts at Root
        '                    '
        '                Else
        '                    Call SeparateURL(Link, Protocol, Host, Path, Page, QueryString)
        '                    If VirtualHosted Then
        '                        '
        '                        ' Virtual hosted site, add VirualPath if it is not there
        '                        '
        '                        If InStr(1, Path, AppRootPath, vbTextCompare) = 0 Then
        '                            If Path = "/" Then
        '                                Path = AppRootPath
        '                            Else
        '                                Path = AppRootPath & Mid(Path, 2)
        '                            End If
        '                        End If
        '                    Else
        '                        '
        '                        ' Root hosted site, remove virtual path if it is there
        '                        '
        '                        If InStr(1, Path, AppRootPath, vbTextCompare) <> 0 Then
        '                            Path = Replace(Path, AppRootPath, "/")
        '                        End If
        '                    End If
        '                    kmaEncodeAppRootPath = Protocol & Host & Path & Page & QueryString
        '                End If
        '            End If
        '        End Function
        '        '
        '        ' Return just the tablename from a tablename reference (database.object.tablename->tablename)
        '        '
        '        Function GetDbObjectTableName(DbObject As String) As String
        '            Dim Position as integer
        '            '
        '            GetDbObjectTableName = DbObject
        '            Position = InStrRev(GetDbObjectTableName, ".")
        '            If Position > 0 Then
        '                GetDbObjectTableName = Mid(GetDbObjectTableName, Position + 1)
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        '        Function kmaGetLinkedText(AnchorTag As Object, AnchorText As Object) As String
        '            '
        '            Dim UcaseAnchorText As String
        '            Dim LinkPosition as integer
        '            Dim MethodName As String
        '            Dim iAnchorTag As String
        '            Dim iAnchorText As String
        '            '
        '            MethodName = "kmaGetLinkedText"
        '            '
        '            kmaGetLinkedText = ""
        '            iAnchorTag = kmaEncodeText(AnchorTag)
        '            iAnchorText = kmaEncodeText(AnchorText)
        '            UcaseAnchorText = UCase(iAnchorText)
        '            If (iAnchorTag <> "") And (iAnchorText <> "") Then
        '                LinkPosition = InStrRev(UcaseAnchorText, "<LINK>", -1)
        '                If LinkPosition = 0 Then
        '                    kmaGetLinkedText = iAnchorTag & iAnchorText & "</A>"
        '                Else
        '                    kmaGetLinkedText = iAnchorText
        '                    LinkPosition = InStrRev(UcaseAnchorText, "</LINK>", -1)
        '                    Do While LinkPosition > 1
        '                        kmaGetLinkedText = Mid(kmaGetLinkedText, 1, LinkPosition - 1) & "</A>" & Mid(kmaGetLinkedText, LinkPosition + 7)
        '                        LinkPosition = InStrRev(UcaseAnchorText, "<LINK>", LinkPosition - 1)
        '                        If LinkPosition <> 0 Then
        '                            kmaGetLinkedText = Mid(kmaGetLinkedText, 1, LinkPosition - 1) & iAnchorTag & Mid(kmaGetLinkedText, LinkPosition + 6)
        '                        End If
        '                        LinkPosition = InStrRev(UcaseAnchorText, "</LINK>", LinkPosition)
        '                    Loop
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        '   HandleError
        '        '       Logs the error and either resumes next, or raises it to the next level
        '        '========================================================================
        '        '
        'Public Function HandleError(ClassName As String, MethodName As String, ErrNumber as integer, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNext As Boolean, Optional URL As String) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim ErrorMessage As String
        '            '
        '            If ErrorTrap Then
        '                ErrorMessage = ErrorMessage & " Unexpected ErrorTrap"
        '            Else
        '                ErrorMessage = ErrorMessage & " Error"
        '            End If
        '            '
        '            If URL <> "" Then
        '                ErrorMessage = ErrorMessage & " on page [" & URL & "]"
        '            End If
        '            '
        '            If ErrorTrap Then
        '                If ResumeNext Then
        '                    Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        '                Else
        '                    Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        '                    On Error GoTo 0
        '                    Call Err.Raise(ErrNumber, ErrSource, ErrDescription)
        '                End If
        '            Else
        '                If ResumeNext Then
        '                    Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging  [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        '                Else
        '                    Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        '                    On Error GoTo 0
        '                    Call Err.Raise(ErrNumber, ErrSource, ErrDescription, , -1)
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        '==========================================================================================
        '        ' handle error and resume next
        '        '==========================================================================================
        '        '
        'Public Sub HandleErrorAndResumeNext(ClassName As String, MethodName As String, Optional Description As String, Optional ErrorNumber as integer)
        '            Dim ErrDescription As String
        '            Dim ErrSource As String
        '            Dim ErrNumber as integer
        '            '
        '            ErrNumber = Err.Number
        '            ErrSource = Err.Source
        '            ErrDescription = Err.Description
        '            '
        '            If ErrNumber = 0 Then
        '                '
        '                ' internal error, no VB error
        '                '
        '                If Description = "" Then
        '                    ErrDescription = "Unknown error"
        '                Else
        '                    ErrDescription = Description
        '                End If
        '                If ErrorNumber = 0 Then
        '                    Call HandleError(ClassName, MethodName, KmaErrorInternal, App.EXEName, Description, False, True)
        '                Else
        '                    Call HandleError(ClassName, MethodName, ErrNumber, App.EXEName, Description, False, True)
        '                End If
        '            Else
        '                '
        '                ' VB Error
        '                '
        '                If Description <> "" Then
        '                    ErrDescription = Description & ",VB Error [" & Err.Description & "]"
        '                Else
        '                    ErrDescription = "Unexpected VB Error [" & Err.Description & "]"
        '                End If
        '                Call HandleError(ClassName, MethodName, ErrNumber, ErrSource, ErrDescription, True, True)
        '            End If
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub AppendLogFile(Text)
        '            On Error GoTo 0
        '            '
        '            Dim MonthNumber as integer
        '            Dim DayNumber as integer
        '            Dim Filename As String
        '            '
        '            DayNumber = Day(Now)
        '            MonthNumber = Month(Now)
        '            Filename = Year(Now)
        '            If MonthNumber < 10 Then
        '                Filename = Filename & "0"
        '            End If
        '            Filename = Filename & MonthNumber
        '            If DayNumber < 10 Then
        '                Filename = Filename & "0"
        '            End If
        '            Filename = Filename & DayNumber
        '            '
        '            Call AppendLog("Trace" & Filename & ".log", kmaEncodeText(Text))
        '            '
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Sub AppendLog(Filename As String, Text As String)
        '            On Error GoTo 0
        '            Dim kmafs As Object
        '            '
        '            If (Filename <> "") Then
        '                kmafs = CreateObject("kmaFileSystem3.FileSystemClass")
        '                Call kmafs.AppendFile(App.Path & "\logs\" & Filename, """" & FormatDateTime(Now(), vbGeneralDate) & """,""" & Text & """" & vbCrLf)
        '                kmafs = Nothing
        '            End If
        '            '
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Function kmaErrorMsg(ErrorNumber as integer) As String
        '            kmaErrorMsg = ""
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaEncodeInitialCaps(Source As String) As String
        '            Dim SegSplit() As String
        '            Dim SegPtr as integer
        '            Dim SegMax as integer
        '            '
        '            If Source <> "" Then
        '                SegSplit = Split(Source, " ")
        '                SegMax = UBound(SegSplit)
        '                If SegMax >= 0 Then
        '                    For SegPtr = 0 To SegMax
        '                        SegSplit(SegPtr) = UCase(Left(SegSplit(SegPtr), 1)) & LCase(Mid(SegSplit(SegPtr), 2))
        '                    Next
        '                End If
        '                kmaEncodeInitialCaps = Join(SegSplit, " ")
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaRemoveTag(Source As String, TagName As String) As String
        '            Dim Pos as integer
        '            Dim posEnd as integer
        '            kmaRemoveTag = Source
        '            Pos = InStr(1, Source, "<" & TagName, vbTextCompare)
        '            If Pos <> 0 Then
        '                posEnd = InStr(Pos, Source, ">")
        '                If posEnd > 0 Then
        '                    kmaRemoveTag = Left(Source, Pos - 1) & Mid(Source, posEnd + 1)
        '                End If
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaRemoveStyleTags(Source As String) As String
        '            kmaRemoveStyleTags = Source
        '            Do While InStr(1, kmaRemoveStyleTags, "<style", vbTextCompare) <> 0
        '                kmaRemoveStyleTags = kmaRemoveTag(kmaRemoveStyleTags, "style")
        '            Loop
        '            Do While InStr(1, kmaRemoveStyleTags, "</style", vbTextCompare) <> 0
        '                kmaRemoveStyleTags = kmaRemoveTag(kmaRemoveStyleTags, "/style")
        '            Loop
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaGetSingular(PluralSource As String) As String
        '            '
        '            Dim UpperCase As Boolean
        '            Dim LastCharacter As String
        '            '
        '            kmaGetSingular = PluralSource
        '            If Len(kmaGetSingular) > 1 Then
        '                LastCharacter = Right(kmaGetSingular, 1)
        '                If LastCharacter <> UCase(LastCharacter) Then
        '                    UpperCase = True
        '                End If
        '                If UCase(Right(kmaGetSingular, 3)) = "IES" Then
        '                    If UpperCase Then
        '                        kmaGetSingular = Mid(kmaGetSingular, 1, Len(kmaGetSingular) - 3) & "Y"
        '                    Else
        '                        kmaGetSingular = Mid(kmaGetSingular, 1, Len(kmaGetSingular) - 3) & "y"
        '                    End If
        '                ElseIf UCase(Right(kmaGetSingular, 2)) = "SS" Then
        '                    ' nothing
        '                ElseIf UCase(Right(kmaGetSingular, 1)) = "S" Then
        '                    kmaGetSingular = Mid(kmaGetSingular, 1, Len(kmaGetSingular) - 1)
        '                Else
        '                    ' nothing
        '                End If
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaEncodeJavascript(Source As String) As String
        '            '
        '            kmaEncodeJavascript = Source
        '            kmaEncodeJavascript = Replace(kmaEncodeJavascript, "'", "\'")
        '            'kmaEncodeJavascript = Replace(kmaEncodeJavascript, "'", "'+""'""+'")
        '            kmaEncodeJavascript = Replace(kmaEncodeJavascript, vbCrLf, "\n")
        '            kmaEncodeJavascript = Replace(kmaEncodeJavascript, vbCr, "\n")
        '            kmaEncodeJavascript = Replace(kmaEncodeJavascript, vbLf, "\n")
        '            kmaEncodeJavascript = Replace(kmaEncodeJavascript, "</script", "</scr'+'ipt", 1, 99, vbTextCompare)
        '            '
        '        End Function
        '
        '   Indent every line by 1 tab
        '
        Public Function kmaIndent(Source As String) As String
            Dim posStart As Integer
            Dim posEnd As Integer
            Dim pre As String
            Dim post As String
            Dim target As String
            '
            posStart = InStr(1, Source, "<![CDATA[", CompareMethod.Text)
            If posStart = 0 Then
                '
                ' no cdata
                '
                posStart = InStr(1, Source, "<textarea", CompareMethod.Text)
                If posStart = 0 Then
                    '
                    ' no textarea
                    '
                    kmaIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
                Else
                    '
                    ' text area found, isolate it and indent before and after
                    '
                    posEnd = InStr(posStart, Source, "</textarea>", CompareMethod.Text)
                    pre = Mid(Source, 1, posStart - 1)
                    If posEnd = 0 Then
                        target = Mid(Source, posStart)
                        post = ""
                    Else
                        target = Mid(Source, posStart, posEnd - posStart + Len("</textarea>"))
                        post = Mid(Source, posEnd + Len("</textarea>"))
                    End If
                    kmaIndent = kmaIndent(pre) & target & kmaIndent(post)
                End If
            Else
                '
                ' cdata found, isolate it and indent before and after
                '
                posEnd = InStr(posStart, Source, "]]>", CompareMethod.Text)
                pre = Mid(Source, 1, posStart - 1)
                If posEnd = 0 Then
                    target = Mid(Source, posStart)
                    post = ""
                Else
                    target = Mid(Source, posStart, posEnd - posStart + Len("]]>"))
                    post = Mid(Source, posEnd + 3)
                End If
                kmaIndent = kmaIndent(pre) & target & kmaIndent(post)
            End If
            '    kmaIndent = Source
            '    If InStr(1, kmaIndent, "<textarea", vbTextCompare) = 0 Then
            '        kmaIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
            '    End If
        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaGetListIndex(Item As String, ListOfItems As String) as integer
        '            '
        '            Dim Items() As String
        '            Dim LcaseItem As String
        '            Dim LcaseList As String
        '            Dim Ptr as integer
        '            '
        '            If ListOfItems <> "" Then
        '                LcaseItem = LCase(Item)
        '                LcaseList = LCase(ListOfItems)
        '                Items = Split(LcaseList, ",")
        '                For Ptr = 0 To UBound(Items)
        '                    If Items(Ptr) = LcaseItem Then
        '                        kmaGetListIndex = Ptr + 1
        '                        Exit For
        '                    End If
        '                Next
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================================================
        '        '
        '        ' Finds all tags matching the input, and concatinates them into the output
        '        ' does NOT account for nested tags, use for body, script, style
        '        '
        '        ' ReturnAll - if true, it returns all the occurances, back-to-back
        '        '
        '        '========================================================================================================
        '        '
        '        Public Function GetTagInnerHTML(PageSource As String, Tag As String, ReturnAll As Boolean) As String
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim TagStart as integer
        '            Dim TagEnd as integer
        '            Dim LoopCnt as integer
        '            Dim WB As String
        '            Dim Pos as integer
        '            Dim posEnd as integer
        '            Dim CommentPos as integer
        '            Dim ScriptPos as integer
        '            '
        '            Pos = 1
        '            Do While (Pos > 0) And (LoopCnt < 100)
        '                TagStart = InStr(Pos, PageSource, "<" & Tag, vbTextCompare)
        '                If TagStart = 0 Then
        '                    Pos = 0
        '                Else
        '                    '
        '                    ' tag found, skip any comments that start between current position and the tag
        '                    '
        '                    CommentPos = InStr(Pos, PageSource, "<!--")
        '                    If (CommentPos <> 0) And (CommentPos < TagStart) Then
        '                        '
        '                        ' skip comment and start again
        '                        '
        '                        Pos = InStr(CommentPos, PageSource, "-->")
        '                    Else
        '                        ScriptPos = InStr(Pos, PageSource, "<script")
        '                        If (ScriptPos <> 0) And (ScriptPos < TagStart) Then
        '                            '
        '                            ' skip comment and start again
        '                            '
        '                            Pos = InStr(ScriptPos, PageSource, "</script")
        '                        Else
        '                            '
        '                            ' Get the tags innerHTML
        '                            '
        '                            TagStart = InStr(TagStart, PageSource, ">", vbTextCompare)
        '                            Pos = TagStart
        '                            If TagStart <> 0 Then
        '                                TagStart = TagStart + 1
        '                                TagEnd = InStr(TagStart, PageSource, "</" & Tag, vbTextCompare)
        '                                If TagEnd <> 0 Then
        '                                    GetTagInnerHTML = GetTagInnerHTML & Mid(PageSource, TagStart, TagEnd - TagStart)
        '                                End If
        '                            End If
        '                        End If
        '                    End If
        '                    LoopCnt = LoopCnt + 1
        '                    If ReturnAll Then
        '                        TagStart = InStr(TagEnd, PageSource, "<" & Tag, vbTextCompare)
        '                    Else
        '                        TagStart = 0
        '                    End If
        '                End If
        '            Loop
        '            '
        '            Exit Function
        '            '
        'ErrorTrap:
        '            'Call HandleError("EncodePage_SplitBody")
        '        End Function
        '        '
        '        '========================================================================================================
        '        'Place code in a form module
        '        'Add a Command button.
        '        '========================================================================================================
        '        '
        '        Public Function kmaByteArrayToString(Bytes() As Byte) As String
        '            Dim iUnicode as integer, i as integer, j as integer

        '            On Error Resume Next
        '            i = UBound(Bytes)

        '            If (i < 1) Then
        '                'ANSI, just convert to unicode and return
        '                kmaByteArrayToString = StrConv(Bytes, vbUnicode)
        '                Exit Function
        '            End If
        '            i = i + 1

        '            'Examine the first two bytes
        '            CopyMemory(iUnicode, Bytes(0), 2)

        '            If iUnicode = Bytes(0) Then 'Unicode
        '                'Account for terminating null
        '                If (i Mod 2) Then i = i - 1
        '                'Set up a buffer to recieve the string
        '                kmaByteArrayToString = String$(i / 2, 0)
        '                'Copy to string
        '        CopyMemory ByVal StrPtr(kmaByteArrayToString), Bytes(0), i
        '            Else 'ANSI
        '                kmaByteArrayToString = StrConv(Bytes, vbUnicode)
        '            End If

        '        End Function
        '        '
        '        '========================================================================================================
        '        '
        '        '========================================================================================================
        '        '
        '        Public Function kmaStringToByteArray(strInput As String, _
        '                                        Optional bReturnAsUnicode As Boolean = True, _
        '                                        Optional bAddNullTerminator As Boolean = False) As Byte()

        '            Dim lRet as integer
        '            Dim bytBuffer() As Byte
        '            Dim lLenB as integer

        '            If bReturnAsUnicode Then
        '                'Number of bytes
        '                lLenB = LenB(strInput)
        '                'Resize buffer, do we want terminating null?
        '                If bAddNullTerminator Then
        '                    ReDim bytBuffer(lLenB)
        '                Else
        '                    ReDim bytBuffer(lLenB - 1)
        '                End If
        '                'Copy characters from string to byte array
        '        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
        '            Else
        '                'METHOD ONE
        '                '        'Get rid of embedded nulls
        '                '        strRet = StrConv(strInput, vbFromUnicode)
        '                '        lLenB = LenB(strRet)
        '                '        If bAddNullTerminator Then
        '                '            ReDim bytBuffer(lLenB)
        '                '        Else
        '                '            ReDim bytBuffer(lLenB - 1)
        '                '        End If
        '                '        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB

        '                'METHOD TWO
        '                'Num of characters
        '                lLenB = Len(strInput)
        '                If bAddNullTerminator Then
        '                    ReDim bytBuffer(lLenB)
        '                Else
        '                    ReDim bytBuffer(lLenB - 1)
        '                End If
        '        lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
        '            End If

        '            kmaStringToByteArray = bytBuffer

        '        End Function
        '        '
        '        '========================================================================================================
        '        '   Sample kmaStringToByteArray
        '        '========================================================================================================
        '        '
        '        Private Sub SampleStringToByteArray()
        '            Dim bAnsi() As Byte
        '            Dim bUni() As Byte
        '            Dim str As String
        '            Dim i as integer
        '            '
        '            str = "Convert"
        '            bAnsi = kmaStringToByteArray(str, False)
        '            bUni = kmaStringToByteArray(str)
        '            '
        '            For i = 0 To UBound(bAnsi)
        '                Debug.Print("=" & bAnsi(i))
        '            Next
        '            '
        '            Debug.Print("========")
        '            '
        '            For i = 0 To UBound(bUni)
        '                Debug.Print("=" & bUni(i))
        '            Next
        '            '
        '            Debug.Print("ANSI= " & kmaByteArrayToString(bAnsi))
        '            Debug.Print("UNICODE= " & kmaByteArrayToString(bUni))
        '            'Using StrConv to convert a Unicode character array directly
        '            'will cause the resultant string to have extra embedded nulls
        '            'reason, StrConv does not know the difference between Unicode and ANSI
        '            Debug.Print("Resull= " & StrConv(bUni, vbUnicode))
        '        End Sub
        '        '
        '        ' ccCommonModule
        '        '
        '        '
        '        '=======================================================================
        '        '   sitepropertyNames
        '        '=======================================================================
        '        '
        '        Public Const siteproperty_serverPageDefault_name = "serverPageDefault"
        '        Public Const siteproperty_serverPageDefault_defaultValue = "index.php"
        '        '
        '        '=======================================================================
        '        '   content replacements
        '        '=======================================================================
        '        '
        '        Public Const contentReplaceEscapeStart = "{%"
        '        Public Const contentReplaceEscapeEnd = "%}"
        '        '
        'Public Type fieldEditorType
        '    fieldId as integer
        '    addonid as integer
        'End Type
        '        '
        '        Public Const protectedContentSetControlFieldList = "ID,CREATEDBY,DATEADDED,MODIFIEDBY,MODIFIEDDATE,EDITSOURCEID,EDITARCHIVE,EDITBLANK,CONTENTCONTROLID"
        '        'Public Const protectedContentSetControlFieldList = "ID,CREATEDBY,DATEADDED,MODIFIEDBY,MODIFIEDDATE,EDITSOURCEID,EDITARCHIVE,EDITBLANK"
        '        '
        '        Public Const HTMLEditorDefaultCopyStartMark = "<!-- cc -->"
        '        Public Const HTMLEditorDefaultCopyEndMark = "<!-- /cc -->"
        '        Public Const HTMLEditorDefaultCopyNoCr = HTMLEditorDefaultCopyStartMark & "<p><br></p>" & HTMLEditorDefaultCopyEndMark
        '        Public Const HTMLEditorDefaultCopyNoCr2 = "<p><br></p>"
        '        '
        '        Public Const IconWidthHeight = " width=21 height=22 "
        '        'Public Const IconWidthHeight = " width=18 height=22 "
        '        '
        '        Public Const CoreCollectionGuid = "{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}" ' contains core Contensive objects, loaded from Library
        '        Public Const ApplicationCollectionGuid = "{C58A76E2-248B-4DE8-BF9C-849A960F79C6}" ' exported from application during upgrade
        '        '
        '        Public Const adminCommonAddonGuid = "{76E7F79E-489F-4B0F-8EE5-0BAC3E4CD782}"
        '        Public Const DashboardAddonGuid = "{4BA7B4A2-ED6C-46C5-9C7B-8CE251FC8FF5}"
        '        Public Const PersonalizationGuid = "{C82CB8A6-D7B9-4288-97FF-934080F5FC9C}"
        '        Public Const TextBoxGuid = "{7010002E-5371-41F7-9C77-0BBFF1F8B728}"
        '        Public Const ContentBoxGuid = "{E341695F-C444-4E10-9295-9BEEC41874D8}"
        '        Public Const DynamicMenuGuid = "{DB1821B3-F6E4-4766-A46E-48CA6C9E4C6E}"
        '        Public Const ChildListGuid = "{D291F133-AB50-4640-9A9A-18DB68FF363B}"
        '        Public Const DynamicFormGuid = "{8284FA0C-6C9D-43E1-9E57-8E9DD35D2DCC}"
        Public Const AddonManagerGuid = "{1DC06F61-1837-419B-AF36-D5CC41E1C9FD}"
        '        Public Const FormWizardGuid = "{2B1384C4-FD0E-4893-B3EA-11C48429382F}"
        '        Public Const ImportWizardGuid = "{37F66F90-C0E0-4EAF-84B1-53E90A5B3B3F}"
        '        Public Const JQueryGuid = "{9C882078-0DAC-48E3-AD4B-CF2AA230DF80}"
        '        Public Const JQueryUIGuid = "{840B9AEF-9470-4599-BD47-7EC0C9298614}"
        '        Public Const ImportProcessAddonGuid = "{5254FAC6-A7A6-4199-8599-0777CC014A13}"
        '        Public Const StructuredDataProcessorGuid = "{65D58FE9-8B76-4490-A2BE-C863B372A6A4}"
        '        Public Const jQueryFancyBoxGuid = "{24C2DBCF-3D84-44B6-A5F7-C2DE7EFCCE3D}"
        '        '
        '        Public Const DefaultLandingPageGuid = "{925F4A57-32F7-44D9-9027-A91EF966FB0D}"
        '        Public Const DefaultLandingSectionGuid = "{D882ED77-DB8F-4183-B12C-F83BD616E2E1}"
        '        Public Const DefaultTemplateGuid = "{47BE95E4-5D21-42CC-9193-A343241E2513}"
        '        Public Const DefaultDynamicMenuGuid = "{E8D575B9-54AE-4BF9-93B7-C7E7FE6F2DB3}"
        '        '
        '        Public Const fpoContentBox = "{1571E62A-972A-4BFF-A161-5F6075720791}"
        '        '
        '        Public Const sfImageExtList = "jpg,jpeg,gif,png"
        '        '
        '        Public Const PageChildListInstanceID = "{ChildPageList}"
        '        '
        Public Const cr = vbCrLf & vbTab
        Public Const cr2 = cr & vbTab
        Public Const cr3 = cr2 & vbTab
        Public Const cr4 = cr3 & vbTab
        Public Const cr5 = cr4 & vbTab
        Public Const cr6 = cr5 & vbTab
        '        '
        '        Public Const AddonOptionConstructor_BlockNoAjax = "Wrapper=[Default:0|None:-1|ListID(Wrappers)]" & vbCrLf & "css Container id" & vbCrLf & "css Container class"
        '        Public Const AddonOptionConstructor_Block = "Wrapper=[Default:0|None:-1|ListID(Wrappers)]" & vbCrLf & "As Ajax=[If Add-on is Ajax:0|Yes:1]" & vbCrLf & "css Container id" & vbCrLf & "css Container class"
        '        Public Const AddonOptionConstructor_Inline = "As Ajax=[If Add-on is Ajax:0|Yes:1]" & vbCrLf & "css Container id" & vbCrLf & "css Container class"
        '        '
        '        ' Constants used as arguments to SiteBuilderClass.CreateNewSite
        '        '
        '        Public Const SiteTypeBaseAsp = 1
        '        Public Const sitetypebaseaspx = 2
        '        Public Const SiteTypeDemoAsp = 3
        '        Public Const SiteTypeBasePhp = 4
        '        '
        '        'Public Const AddonNewParse = True
        '        '
        '        Public Const AddonOptionConstructor_ForBlockText = "AllowGroups=[listid(groups)]checkbox"
        '        Public Const AddonOptionConstructor_ForBlockTextEnd = ""
        '        Public Const BlockTextStartMarker = "<!-- BLOCKTEXTSTART -->"
        '        Public Const BlockTextEndMarker = "<!-- BLOCKTEXTEND -->"
        '        '
        '        Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as integer)
        '        Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess as integer, lpExitCode as integer) as integer
        '        Private Declare Function timeGetTime Lib "winmm.dll" () as integer
        '        Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess as integer, ByVal bInheritHandle as integer, ByVal dwProcessId as integer) as integer
        '        Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject as integer) as integer
        '        '
        '        Public Const InstallFolderName = "Install"
        '        Public Const DownloadFileRootNode = "collectiondownload"
        '        Public Const CollectionFileRootNode = "collection"
        '        Public Const CollectionFileRootNodeOld = "addoncollection"
        '        Public Const CollectionListRootNode = "collectionlist"
        '        '
        '        Public Const LegacyLandingPageName = "Landing Page Content"
        '        Public Const DefaultNewLandingPageName = "Home"
        '        Public Const DefaultLandingSectionName = "Home"
        '        '
        '        ' ----- Errors Specific to the Contensive Objects
        '        '
        '        Public Const KmaccErrorUpgrading = KmaObjectError + 1
        '        Public Const KmaccErrorServiceStopped = KmaObjectError + 2
        '        '
        '        Public Const UserErrorHeadline = "<p class=""ccError"">There was a problem with this page.</p>"
        '        '
        '        ' ----- Errors connecting to server
        '        '
        '        Public Const ccError_InvalidAppName = 100
        '        Public Const ccError_ErrorAddingApp = 101
        '        Public Const ccError_ErrorDeletingApp = 102
        '        Public Const ccError_InvalidFieldName = 103     ' Invalid parameter name
        '        Public Const ccError_InvalidCommand = 104
        '        Public Const ccError_InvalidAuthentication = 105
        '        Public Const ccError_NotConnected = 106             ' Attempt to execute a command without a connection
        '        '
        '        '
        '        '
        '        Public Const ccStatusCode_Base = KmaErrorBase
        '        Public Const ccStatusCode_ControllerCreateFailed = ccStatusCode_Base + 1
        '        Public Const ccStatusCode_ControllerInProcess = ccStatusCode_Base + 2
        '        Public Const ccStatusCode_ControllerStartedWithoutService = ccStatusCode_Base + 3
        '        '
        '        ' ----- Previous errors, can be replaced
        '        '
        '        'Public Const KmaError_UnderlyingObject_Msg = "An error occurred in an underlying routine."
        '        Public Const KmaccErrorServiceStopped_Msg = "The Contensive CSv Service is not running."
        '        Public Const KmaError_BadObject_Msg = "Server Object is not valid."
        '        Public Const KmaError_UpgradeInProgress_Msg = "Server is busy with internal upgrade."
        '        '
        '        'Public Const KmaError_InvalidArgument_Msg = "Invalid Argument"
        '        'Public Const KmaError_UnderlyingObject_Msg = "An error occurred in an underlying routine."
        '        'Public Const KmaccErrorServiceStopped_Msg = "The Contensive CSv Service is not running."
        '        'Public Const KmaError_BadObject_Msg = "Server Object is not valid."
        '        'Public Const KmaError_UpgradeInProgress_Msg = "Server is busy with internal upgrade."
        '        'Public Const KmaError_InvalidArgument_Msg = "Invalid Argument"
        '        '
        '        '-----------------------------------------------------------------------
        '        '   GetApplicationList indexes
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const AppList_Name = 0
        '        Public Const AppList_Status = 1
        '        Public Const AppList_ConnectionsActive = 2
        '        Public Const AppList_ConnectionString = 3
        '        Public Const AppList_DataBuildVersion = 4
        '        Public Const AppList_LicenseKey = 5
        '        Public Const AppList_RootPath = 6
        '        Public Const AppList_PhysicalFilePath = 7
        '        Public Const AppList_DomainName = 8
        '        Public Const AppList_DefaultPage = 9
        '        Public Const AppList_AllowSiteMonitor = 10
        '        Public Const AppList_HitCounter = 11
        '        Public Const AppList_ErrorCount = 12
        '        Public Const AppList_DateStarted = 13
        '        Public Const AppList_AutoStart = 14
        '        Public Const AppList_Progress = 15
        '        Public Const AppList_PhysicalWWWPath = 16
        '        Public Const AppListCount = 17
        '        '
        '        '-----------------------------------------------------------------------
        '        '   System MemberID - when the system does an update, it uses this member
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const SystemMemberID = 0
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- old (OptionKeys for available Options)
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const OptionKeyProductionLicense = 0
        '        Public Const OptionKeyDeveloperLicense = 1
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- LicenseTypes, replaced OptionKeys
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const LicenseTypeInvalid = -1
        '        Public Const LicenseTypeProduction = 0
        '        Public Const LicenseTypeTrial = 1
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- Active Content Definitions
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const ACTypeDate = "DATE"
        '        Public Const ACTypeVisit = "VISIT"
        '        Public Const ACTypeVisitor = "VISITOR"
        '        Public Const ACTypeMember = "MEMBER"
        '        Public Const ACTypeOrganization = "ORGANIZATION"
        '        Public Const ACTypeChildList = "CHILDLIST"
        '        Public Const ACTypeContact = "CONTACT"
        '        Public Const ACTypeFeedback = "FEEDBACK"
        '        Public Const ACTypeLanguage = "LANGUAGE"
        '        Public Const ACTypeAggregateFunction = "AGGREGATEFUNCTION"
        '        Public Const ACTypeAddon = "ADDON"
        '        Public Const ACTypeImage = "IMAGE"
        '        Public Const ACTypeDownload = "DOWNLOAD"
        '        Public Const ACTypeEnd = "END"
        '        Public Const ACTypeTemplateContent = "CONTENT"
        '        Public Const ACTypeTemplateText = "TEXT"
        '        Public Const ACTypeDynamicMenu = "DYNAMICMENU"
        '        Public Const ACTypeWatchList = "WATCHLIST"
        '        Public Const ACTypeRSSLink = "RSSLINK"
        '        Public Const ACTypePersonalization = "PERSONALIZATION"
        '        Public Const ACTypeDynamicForm = "DYNAMICFORM"
        '        '
        '        Public Const ACTagEnd = "<ac type=""" & ACTypeEnd & """>"
        '        '
        '        ' ----- PropertyType Definitions
        '        '
        '        Public Const PropertyTypeMember = 0
        '        Public Const PropertyTypeVisit = 1
        '        Public Const PropertyTypeVisitor = 2
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- Port Assignments
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const WinsockPortWebOut = 4000
        '        Public Const WinsockPortServerFromWeb = 4001
        '        Public Const WinsockPortServerToClient = 4002
        '        '
        '        Public Const Port_ContentServerControlDefault = 4531
        '        Public Const Port_SiteMonitorDefault = 4532
        '        '
        '        Public Const RMBMethodHandShake = 1
        '        Public Const RMBMethodMessage = 3
        '        Public Const RMBMethodTestPoint = 4
        '        Public Const RMBMethodInit = 5
        '        Public Const RMBMethodClosePage = 6
        '        Public Const RMBMethodOpenCSContent = 7
        '        '
        '        ' ----- Position equates for the Remote Method Block
        '        '
        '        Const RMBPositionLength = 0             ' Length of the RMB
        '        Const RMBPositionSourceHandle = 4       ' Handle generated by the source of the command
        '        Const RMBPositionMethod = 8             ' Method in the method block
        '        Const RMBPositionArgumentCount = 12     ' The number of arguments in the Block
        '        Const RMBPositionFirstArgument = 16     ' The offset to the first argu
        '        '
        '        '-----------------------------------------------------------------------
        '        '   Remote Connections
        '        '   List of current remove connections for Remote Monitoring/administration
        '        '-----------------------------------------------------------------------
        '        '
        'Public Type RemoteAdministratorType
        '    RemoteIP As String
        '    RemotePort as integer
        'End Type
        '        '
        '        ' Default username/password
        '        '
        '        Public Const DefaultServerUsername = "root"
        '        Public Const DefaultServerPassword = "contensive"
        '        '
        '        '-----------------------------------------------------------------------
        '        '   Form Contension Strategy
        '        '
        '        '       all Contensive Forms contain a hidden "ccFormSN"
        '        '       The value in the hidden is the FormID string. All input
        '        '       elements of the form are named FormID & "ElementName"
        '        '
        '        '       This prevents form elements from different forms from interfearing
        '        '       with each other, and with developer generated forms.
        '        '
        '        '       GetFormSN gets a new valid random formid to be used.
        '        '       All forms requires:
        '        '           a FormId (text), containing the formid string
        '        '           a [formid]Type (text), as defined in FormTypexxx in CommonModule
        '        '
        '        '       Forms have two primary sections: GetForm and ProcessForm
        '        '
        '        '       Any form that has a GetForm method, should have the process form
        '        '       in the main.init, selected with this [formid]type hidden (not the
        '        '       GetForm method). This is so the process can alter the stream
        '        '       output for areas before the GetForm call.
        '        '
        '        '       System forms, like tools panel, that may appear on any page, have
        '        '       their process call in the main.init.
        '        '
        '        '       Popup forms, like ImageSelector have their processform call in the
        '        '       main.init because no .asp page exists that might contain a call
        '        '       the process section.
        '        '
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const FormTypeToolsPanel = "do30a8vl29"
        '        Public Const FormTypeActiveEditor = "l1gk70al9n"
        '        Public Const FormTypeImageSelector = "ila9c5s01m"
        '        Public Const FormTypePageAuthoring = "2s09lmpalb"
        '        Public Const FormTypeMyProfile = "89aLi180j5"
        '        Public Const FormTypeLogin = "login"
        '        'Public Const FormTypeLogin = "l09H58a195"
        '        Public Const FormTypeSendPassword = "lk0q56am09"
        '        Public Const FormTypeJoin = "6df38abv00"
        '        Public Const FormTypeHelpBubbleEditor = "9df019d77sA"
        '        Public Const FormTypeAddonSettingsEditor = "4ed923aFGw9d"
        '        Public Const FormTypeAddonStyleEditor = "ar5028jklkfd0s"
        '        Public Const FormTypeSiteStyleEditor = "fjkq4w8794kdvse"
        '        'Public Const FormTypeAggregateFunctionProperties = "9wI751270"
        '        '
        '        '-----------------------------------------------------------------------
        '        '   Hardcoded profile form const
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const rnMyProfileTopics = "profileTopics"
        '        '
        '        '-----------------------------------------------------------------------
        '        ' Legacy - replaced with HardCodedPages
        '        '   Intercept Page Strategy
        '        '
        '        '       RequestnameInterceptpage = InterceptPage number from the input stream
        '        '       InterceptPage = Global variant with RequestnameInterceptpage value read during early Init
        '        '
        '        '       Intercept pages are complete pages that appear instead of what
        '        '       the physical page calls.
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const RequestNameInterceptpage = "ccIPage"
        '        '
        '        Public Const LegacyInterceptPageSNResourceLibrary = "s033l8dm15"
        '        Public Const LegacyInterceptPageSNSiteExplorer = "kdif3318sd"
        '        Public Const LegacyInterceptPageSNImageUpload = "ka983lm039"
        '        Public Const LegacyInterceptPageSNMyProfile = "k09ddk9105"
        '        Public Const LegacyInterceptPageSNLogin = "6ge42an09a"
        '        Public Const LegacyInterceptPageSNPrinterVersion = "l6d09a10sP"
        '        Public Const LegacyInterceptPageSNUploadEditor = "k0hxp2aiOZ"
        '        '
        '        '-----------------------------------------------------------------------
        '        ' Ajax functions intercepted during init, answered and response closed
        '        '   These are hard-coded internal Contensive functions
        '        '   These should eventually be replaced with (HardcodedAddons) remote methods
        '        '   They should all be prefixed "cc"
        '        '   They are called with cj.ajax.qs(), setting RequestNameAjaxFunction=name in the qs
        '        '   These name=value pairs go in the QueryString argument of the javascript cj.ajax.qs() function
        '        '-----------------------------------------------------------------------
        '        '
        '        'Public Const RequestNameOpenSettingPage = "settingpageid"
        '        Public Const RequestNameAjaxFunction = "ajaxfn"
        '        Public Const RequestNameAjaxFastFunction = "ajaxfastfn"
        '        '
        '        Public Const AjaxOpenAdminNav = "aps89102kd"
        '        Public Const AjaxOpenAdminNavGetContent = "d8475jkdmfj2"
        '        Public Const AjaxCloseAdminNav = "3857fdjdskf91"
        '        Public Const AjaxAdminNavOpenNode = "8395j2hf6jdjf"
        '        Public Const AjaxAdminNavOpenNodeGetContent = "eieofdwl34efvclaeoi234598"
        '        Public Const AjaxAdminNavCloseNode = "w325gfd73fhdf4rgcvjk2"
        '        '
        '        Public Const AjaxCloseIndexFilter = "k48smckdhorle0"
        '        Public Const AjaxOpenIndexFilter = "Ls8jCDt87kpU45YH"
        '        Public Const AjaxOpenIndexFilterGetContent = "llL98bbJQ38JC0KJm"
        '        Public Const AjaxStyleEditorAddStyle = "ajaxstyleeditoradd"
        '        Public Const AjaxPing = "ajaxalive"
        '        Public Const AjaxGetFormEditTabContent = "ajaxgetformedittabcontent"
        '        Public Const AjaxData = "data"
        '        Public Const AjaxGetVisitProperty = "getvisitproperty"
        '        Public Const AjaxSetVisitProperty = "setvisitproperty"
        '        Public Const AjaxGetDefaultAddonOptionString = "ccGetDefaultAddonOptionString"
        '        Public Const ajaxGetFieldEditorPreferenceForm = "ajaxgetfieldeditorpreference"
        '        '
        '        '-----------------------------------------------------------------------
        '        '
        '        ' no - for now just use ajaxfn in the cj.ajax.qs call
        '        '   this is more work, and I do not see why it buys anything new or better
        '        '
        '        '   Hard-coded addons
        '        '       these are internal Contensive functions
        '        '       can be called with just /addonname?querystring
        '        '       call them with cj.ajax.addon() or cj.ajax.addonCallback()
        '        '       are first in the list of checks when a URL rewrite is detected in Init()
        '        '       should all be prefixed with 'cc'
        '        '-----------------------------------------------------------------------
        '        '
        '        'Public Const HardcodedAddonGetDefaultAddonOptionString = "ccGetDefaultAddonOptionString"
        '        '
        '        '-----------------------------------------------------------------------
        '        '   Remote Methods
        '        '       ?RemoteMethodAddon=string
        '        '       calls an addon (if marked to run as a remote method)
        '        '       blocks all other Contensive output (tools panel, javascript, etc)
        '        '-----------------------------------------------------------------------
        '        '
        Public Const RequestNameRemoteMethodAddon = "remotemethodaddon"
        '        '
        '        '-----------------------------------------------------------------------
        '        '   Hard Coded Pages
        '        '       ?Method=string
        '        '       Querystring based so they can be added to URLs, preserving the current page for a return
        '        '       replaces output stream with html output
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const RequestNameHardCodedPage = "method"
        '        '
        '        Public Const HardCodedPageLogin = "login"
        '        Public Const HardCodedPageLoginDefault = "logindefault"
        '        Public Const HardCodedPageMyProfile = "myprofile"
        '        Public Const HardCodedPagePrinterVersion = "printerversion"
        '        Public Const HardCodedPageResourceLibrary = "resourcelibrary"
        '        Public Const HardCodedPageLogoutLogin = "logoutlogin"
        '        Public Const HardCodedPageLogout = "logout"
        '        Public Const HardCodedPageSiteExplorer = "siteexplorer"
        '        'Public Const HardCodedPageForceMobile = "forcemobile"
        '        'Public Const HardCodedPageForceNonMobile = "forcenonmobile"
        '        Public Const HardCodedPageNewOrder = "neworderpage"
        '        Public Const HardCodedPageStatus = "status"
        '        Public Const HardCodedPageGetJSPage = "getjspage"
        '        Public Const HardCodedPageGetJSLogin = "getjslogin"
        '        Public Const HardCodedPageRedirect = "redirect"
        '        Public Const HardCodedPageExportAscii = "exportascii"
        '        Public Const HardCodedPagePayPalConfirm = "paypalconfirm"
        '        Public Const HardCodedPageSendPassword = "sendpassword"
        '        '
        '        '-----------------------------------------------------------------------
        '        '   Option values
        '        '       does not effect output directly
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const RequestNamePageOptions = "ccoptions"
        '        '
        '        Public Const PageOptionForceMobile = "forcemobile"
        '        Public Const PageOptionForceNonMobile = "forcenonmobile"
        '        Public Const PageOptionLogout = "logout"
        '        Public Const PageOptionPrinterVersion = "printerversion"
        '        '
        '        ' convert to options later
        '        '
        '        Public Const RequestNameDashboardReset = "ResetDashboard"
        '        '
        '        '-----------------------------------------------------------------------
        '        '   DataSource constants
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const DefaultDataSourceID = -1
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- Type compatibility between databases
        '        '       Boolean
        '        '           Access      YesNo       true=1, false=0
        '        '           SQL Server  bit         true=1, false=0
        '        '           MySQL       bit         true=1, false=0
        '        '           Oracle      integer(1)  true=1, false=0
        '        '           Note: false does not equal NOT true
        '        '       Integer (Number)
        '        '           Access      Long        8 bytes, about E308
        '        '           SQL Server  int
        '        '           MySQL       integer
        '        '           Oracle      integer(8)
        '        '       Float
        '        '           Access      Double      8 bytes, about E308
        '        '           SQL Server  Float
        '        '           MySQL
        '        '           Oracle
        '        '       Text
        '        '           Access
        '        '           SQL Server
        '        '           MySQL
        '        '           Oracle
        '        '-----------------------------------------------------------------------
        '        '
        '        'Public Const SQLFalse = "0"
        '        'Public Const SQLTrue = "1"
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- Style sheet definitions
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const defaultStyleFilename = "ccDefault.r5.css"
        '        Public Const StyleSheetStart = "<STYLE TYPE=""text/css"">"
        '        Public Const StyleSheetEnd = "</STYLE>"
        '        '
        '        Public Const SpanClassAdminNormal = "<span class=""ccAdminNormal"">"
        '        Public Const SpanClassAdminSmall = "<span class=""ccAdminSmall"">"
        '        '
        '        ' remove these from ccWebx
        '        '
        '        Public Const SpanClassNormal = "<span class=""ccNormal"">"
        '        Public Const SpanClassSmall = "<span class=""ccSmall"">"
        '        Public Const SpanClassLarge = "<span class=""ccLarge"">"
        '        Public Const SpanClassHeadline = "<span class=""ccHeadline"">"
        '        Public Const SpanClassList = "<span class=""ccList"">"
        '        Public Const SpanClassListCopy = "<span class=""ccListCopy"">"
        '        Public Const SpanClassError = "<span class=""ccError"">"
        '        Public Const SpanClassSeeAlso = "<span class=""ccSeeAlso"">"
        '        Public Const SpanClassEnd = "</span>"
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- XHTML definitions
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const DTDTransitional = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
        '        '
        '        Public Const BR = "<br>"
        '        '
        '        '-----------------------------------------------------------------------
        '        ' AuthoringControl Types
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const AuthoringControlsEditing = 1
        '        Public Const AuthoringControlsSubmitted = 2
        '        Public Const AuthoringControlsApproved = 3
        '        Public Const AuthoringControlsModified = 4
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- Panel and header colors
        '        '-----------------------------------------------------------------------
        '        '
        '        'Public Const "ccPanel" = "#E0E0E0"    ' The background color of a panel (black copy visible on it)
        '        'Public Const "ccPanelHilite" = "#F8F8F8"  '
        '        'Public Const "ccPanelShadow" = "#808080"  '
        '        '
        '        'Public Const HeaderColorBase = "#0320B0"   ' The background color of a panel header (reverse copy visible)
        '        'Public Const "ccPanelHeaderHilite" = "#8080FF" '
        '        'Public Const "ccPanelHeaderShadow" = "#000000" '
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- Field type Definitions
        '        '       Field Types are numeric values that describe how to treat values
        '        '       stored as ContentFieldDefinitionType (FieldType property of FieldType Type.. ;)
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const FieldTypeInteger = 1       ' An long number
        '        Public Const FieldTypeText = 2          ' A text field (up to 255 characters)
        '        Public Const FieldTypeLongText = 3      ' A memo field (up to 8000 characters)
        '        Public Const FieldTypeBoolean = 4       ' A yes/no field
        '        Public Const FieldTypeDate = 5          ' A date field
        '        Public Const FieldTypeFile = 6          ' A filename of a file in the files directory.
        '        Public Const FieldTypeLookup = 7        ' A lookup is a FieldTypeInteger that indexes into another table
        '        Public Const FieldTypeRedirect = 8      ' creates a link to another section
        '        Public Const FieldTypeCurrency = 9      ' A Float that prints in dollars
        '        Public Const FieldTypeTextFile = 10     ' Text saved in a file in the files area.
        '        Public Const FieldTypeImage = 11        ' A filename of a file in the files directory.
        '        Public Const FieldTypeFloat = 12        ' A float number
        '        Public Const FieldTypeAutoIncrement = 13 'long that automatically increments with the new record
        '        Public Const FieldTypeManyToMany = 14    ' no database field - sets up a relationship through a Rule table to another table
        '        Public Const FieldTypeMemberSelect = 15 ' This ID is a ccMembers record in a group defined by the MemberSelectGroupID field
        '        Public Const FieldTypeCSSFile = 16      ' A filename of a CSS compatible file
        '        Public Const FieldTypeXMLFile = 17      ' the filename of an XML compatible file
        '        Public Const FieldTypeJavascriptFile = 18 ' the filename of a javascript compatible file
        '        Public Const FieldTypeLink = 19           ' Links used in href tags -- can go to pages or resources
        '        Public Const FieldTypeResourceLink = 20   ' Links used in resources, link <img or <object. Should not be pages
        '        Public Const FieldTypeHTML = 21           ' LongText field that expects HTML content
        '        Public Const FieldTypeHTMLFile = 22       ' TextFile field that expects HTML content
        '        Public Const FieldTypeMax = 22
        '        '
        '        ' ----- Field Descriptors for these type
        '        '       These are what are publicly displayed for each type
        '        '       See GetFieldDescriptorByType and vise-versa to translater
        '        '
        '        Public Const FieldDescriptorInteger = "Integer"
        '        Public Const FieldDescriptorText = "Text"
        '        Public Const FieldDescriptorLongText = "LongText"
        '        Public Const FieldDescriptorBoolean = "Boolean"
        '        Public Const FieldDescriptorDate = "Date"
        '        Public Const FieldDescriptorFile = "File"
        '        Public Const FieldDescriptorLookup = "Lookup"
        '        Public Const FieldDescriptorRedirect = "Redirect"
        '        Public Const FieldDescriptorCurrency = "Currency"
        '        Public Const FieldDescriptorImage = "Image"
        '        Public Const FieldDescriptorFloat = "Float"
        '        Public Const FieldDescriptorManyToMany = "ManyToMany"
        '        Public Const FieldDescriptorTextFile = "TextFile"
        '        Public Const FieldDescriptorCSSFile = "CSSFile"
        '        Public Const FieldDescriptorXMLFile = "XMLFile"
        '        Public Const FieldDescriptorJavascriptFile = "JavascriptFile"
        '        Public Const FieldDescriptorLink = "Link"
        '        Public Const FieldDescriptorResourceLink = "ResourceLink"
        '        Public Const FieldDescriptorMemberSelect = "MemberSelect"
        '        Public Const FieldDescriptorHTML = "HTML"
        '        Public Const FieldDescriptorHTMLFile = "HTMLFile"
        '        '
        '        Public Const FieldDescriptorLcaseInteger = "integer"
        '        Public Const FieldDescriptorLcaseText = "text"
        '        Public Const FieldDescriptorLcaseLongText = "longtext"
        '        Public Const FieldDescriptorLcaseBoolean = "boolean"
        '        Public Const FieldDescriptorLcaseDate = "date"
        '        Public Const FieldDescriptorLcaseFile = "file"
        '        Public Const FieldDescriptorLcaseLookup = "lookup"
        '        Public Const FieldDescriptorLcaseRedirect = "redirect"
        '        Public Const FieldDescriptorLcaseCurrency = "currency"
        '        Public Const FieldDescriptorLcaseImage = "image"
        '        Public Const FieldDescriptorLcaseFloat = "float"
        '        Public Const FieldDescriptorLcaseManyToMany = "manytomany"
        '        Public Const FieldDescriptorLcaseTextFile = "textfile"
        '        Public Const FieldDescriptorLcaseCSSFile = "cssfile"
        '        Public Const FieldDescriptorLcaseXMLFile = "xmlfile"
        '        Public Const FieldDescriptorLcaseJavascriptFile = "javascriptfile"
        '        Public Const FieldDescriptorLcaseLink = "link"
        '        Public Const FieldDescriptorLcaseResourceLink = "resourcelink"
        '        Public Const FieldDescriptorLcaseMemberSelect = "memberselect"
        '        Public Const FieldDescriptorLcaseHTML = "html"
        '        Public Const FieldDescriptorLcaseHTMLFile = "htmlfile"
        '        '
        '        '------------------------------------------------------------------------
        '        ' ----- Payment Options
        '        '------------------------------------------------------------------------
        '        '
        '        Public Const PayTypeCreditCardOnline = 0   ' Pay by credit card online
        '        Public Const PayTypeCreditCardByPhone = 1  ' Phone in a credit card
        '        Public Const PayTypeCreditCardByFax = 9    ' Phone in a credit card
        '        Public Const PayTypeCHECK = 2              ' pay by check to be mailed
        '        Public Const PayTypeCREDIT = 3             ' pay on account
        '        Public Const PayTypeNONE = 4               ' order total is $0.00. Nothing due
        '        Public Const PayTypeCHECKCOMPANY = 5       ' pay by company check
        '        Public Const PayTypeNetTerms = 6
        '        Public Const PayTypeCODCompanyCheck = 7
        '        Public Const PayTypeCODCertifiedFunds = 8
        '        Public Const PayTypePAYPAL = 10
        '        Public Const PAYDEFAULT = 0
        '        '
        '        '------------------------------------------------------------------------
        '        ' ----- Credit card options
        '        '------------------------------------------------------------------------
        '        '
        '        Public Const CCTYPEVISA = 0                ' Visa
        '        Public Const CCTYPEMC = 1                  ' MasterCard
        '        Public Const CCTYPEAMEX = 2                ' American Express
        '        Public Const CCTYPEDISCOVER = 3            ' Discover
        '        Public Const CCTYPENOVUS = 4               ' Novus Card
        '        Public Const CCTYPEDEFAULT = 0
        '        '
        '        '------------------------------------------------------------------------
        '        ' ----- Shipping Options
        '        '------------------------------------------------------------------------
        '        '
        '        Public Const SHIPGROUND = 0                ' ground
        '        Public Const SHIPOVERNIGHT = 1             ' overnight
        '        Public Const SHIPSTANDARD = 2              ' standard, whatever that is
        '        Public Const SHIPOVERSEAS = 3              ' to overseas
        '        Public Const SHIPCANADA = 4                ' to Canada
        '        Public Const SHIPDEFAULT = 0
        '        '
        '        '------------------------------------------------------------------------
        '        ' Debugging info
        '        '------------------------------------------------------------------------
        '        '
        '        Public Const TestPointTab = 2
        '        Public Const TestPointTabChr = "-"
        '        Public CPTickCountBase As Double
        '        '
        '        '------------------------------------------------------------------------
        '        '   project width button defintions
        '        '------------------------------------------------------------------------
        '        '
        '        Public Const ButtonApply = "  Apply "
        '        Public Const ButtonLogin = "  Login  "
        '        Public Const ButtonLogout = "  Logout  "
        '        Public Const ButtonSendPassword = "  Send Password  "
        '        Public Const ButtonJoin = "   Join   "
        '        Public Const ButtonSave = "  Save  "
        '        Public Const ButtonOK = "     OK     "
        '        Public Const ButtonReset = "  Reset  "
        '        Public Const ButtonSaveAddNew = " Save + Add "
        '        'Public Const ButtonSaveAddNew = " Save > Add "
        '        Public Const ButtonCancel = " Cancel "
        '        Public Const ButtonRestartContensiveApplication = " Restart Contensive Application "
        '        Public Const ButtonCancelAll = "  Cancel  "
        '        Public Const ButtonFind = "   Find   "
        '        Public Const ButtonDelete = "  Delete  "
        '        Public Const ButtonDeletePerson = " Delete Person "
        '        Public Const ButtonDeleteRecord = " Delete Record "
        '        Public Const ButtonDeleteEmail = " Delete Email "
        '        Public Const ButtonDeletePage = " Delete Page "
        '        Public Const ButtonFileChange = "   Upload   "
        '        Public Const ButtonFileDelete = "    Delete    "
        '        Public Const ButtonClose = "  Close   "
        '        Public Const ButtonAdd = "   Add    "
        '        Public Const ButtonAddChildPage = " Add Child "
        '        Public Const ButtonAddSiblingPage = " Add Sibling "
        '        Public Const ButtonContinue = " Continue >> "
        '        Public Const ButtonBack = "  << Back  "
        '        Public Const ButtonNext = "   Next   "
        '        Public Const ButtonPrevious = " Previous "
        '        Public Const ButtonFirst = "  First   "
        '        Public Const ButtonSend = "  Send   "
        '        Public Const ButtonSendTest = "Send Test"
        '        Public Const ButtonCreateDuplicate = " Create Duplicate "
        '        Public Const ButtonActivate = "  Activate   "
        '        Public Const ButtonDeactivate = "  Deactivate   "
        '        Public Const ButtonOpenActiveEditor = "Active Edit"
        '        Public Const ButtonPublish = " Publish Changes "
        '        Public Const ButtonAbortEdit = " Abort Edits "
        '        Public Const ButtonPublishSubmit = " Submit for Publishing "
        '        Public Const ButtonPublishApprove = " Approve for Publishing "
        '        Public Const ButtonPublishDeny = " Deny for Publishing "
        '        Public Const ButtonWorkflowPublishApproved = " Publish Approved Records "
        '        Public Const ButtonWorkflowPublishSelected = " Publish Selected Records "
        '        Public Const ButtonSetHTMLEdit = " Edit WYSIWYG "
        '        Public Const ButtonSetTextEdit = " Edit HTML "
        '        Public Const ButtonRefresh = " Refresh "
        '        Public Const ButtonOrder = " Order "
        '        Public Const ButtonSearch = " Search "
        '        Public Const ButtonSpellCheck = " Spell Check "
        '        Public Const ButtonLibraryUpload = " Upload "
        '        Public Const ButtonCreateReport = " Create Report "
        '        Public Const ButtonClearTrapLog = " Clear Trap Log "
        '        Public Const ButtonNewSearch = " New Search "
        '        Public Const ButtonReloadCDef = " Reload Content Definitions "
        '        Public Const ButtonImportTemplates = " Import Templates "
        '        Public Const ButtonRSSRefresh = " Update RSS Feeds Now "
        '        Public Const ButtonRequestDownload = " Request Download "
        '        Public Const ButtonFinish = " Finish "
        '        Public Const ButtonRegister = " Register "
        '        Public Const ButtonBegin = "Begin"
        '        Public Const ButtonAbort = "Abort"
        '        Public Const ButtonCreateGUID = " Create GUID "
        '        Public Const ButtonEnable = " Enable "
        '        Public Const ButtonDisable = " Disable "
        '        Public Const ButtonMarkReviewed = " Mark Reviewed "
        '        '
        '        '------------------------------------------------------------------------
        '        '   member actions
        '        '------------------------------------------------------------------------
        '        '
        '        Public Const MemberActionNOP = 0
        '        Public Const MemberActionLogin = 1
        '        Public Const MemberActionLogout = 2
        '        Public Const MemberActionForceLogin = 3
        '        Public Const MemberActionSendPassword = 4
        '        Public Const MemberActionForceLogout = 5
        '        Public Const MemberActionToolsApply = 6
        '        Public Const MemberActionJoin = 7
        '        Public Const MemberActionSaveProfile = 8
        '        Public Const MemberActionEditProfile = 9
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- note pad info
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const NoteFormList = 1
        '        Public Const NoteFormRead = 2
        '        '
        '        Public Const NoteButtonPrevious = " Previous "
        '        Public Const NoteButtonNext = "   Next   "
        '        Public Const NoteButtonDelete = "  Delete  "
        '        Public Const NoteButtonClose = "  Close   "
        '        '                       ' Submit button is created in CommonDim, so it is simple
        '        Public Const NoteButtonSubmit = "Submit"
        '        '
        '        '-----------------------------------------------------------------------
        '        ' ----- Admin site storage
        '        '-----------------------------------------------------------------------
        '        '
        '        Public Const AdminMenuModeHidden = 0       '   menu is hidden
        '        Public Const AdminMenuModeLeft = 1     '   menu on the left
        '        Public Const AdminMenuModeTop = 2      '   menu as dropdowns from the top
        '        '
        '        ' ----- AdminActions - describes the form processing to do
        '        '
        '        Public Const AdminActionNop = 0            ' do nothing
        '        Public Const AdminActionDelete = 4         ' delete record
        '        Public Const AdminActionFind = 5           '
        '        Public Const AdminActionDeleteFilex = 6        '
        '        Public Const AdminActionUpload = 7         '
        '        Public Const AdminActionSaveNormal = 3         ' save fields to database
        '        Public Const AdminActionSaveEmail = 8          ' save email record (and update EmailGroups) to database
        '        Public Const AdminActionSaveMember = 11        '
        '        Public Const AdminActionSaveSystem = 12
        '        Public Const AdminActionSavePaths = 13     ' Save a record that is in the BathBlocking Format
        '        Public Const AdminActionSendEmail = 9          '
        '        Public Const AdminActionSendEmailTest = 10     '
        '        Public Const AdminActionNext = 14               '
        '        Public Const AdminActionPrevious = 15           '
        '        Public Const AdminActionFirst = 16              '
        '        Public Const AdminActionSaveContent = 17        '
        '        Public Const AdminActionSaveField = 18          ' Save a single field, fieldname = fn input
        '        Public Const AdminActionPublish = 19            ' Publish record live
        '        Public Const AdminActionAbortEdit = 20          ' Publish record live
        '        Public Const AdminActionPublishSubmit = 21      ' Submit for Workflow Publishing
        '        Public Const AdminActionPublishApprove = 22     ' Approve for Workflow Publishing
        '        Public Const AdminActionWorkflowPublishApproved = 23    ' Publish what was approved
        '        Public Const AdminActionSetHTMLEdit = 24        ' Set Member Property for this field to HTML Edit
        '        Public Const AdminActionSetTextEdit = 25        ' Set Member Property for this field to Text Edit
        '        Public Const AdminActionSave = 26               ' Save Record
        '        Public Const AdminActionActivateEmail = 27      ' Activate a Conditional Email
        '        Public Const AdminActionDeactivateEmail = 28    ' Deactivate a conditional email
        '        Public Const AdminActionDuplicate = 29          ' Duplicate the (sent email) record
        '        Public Const AdminActionDeleteRows = 30         ' Delete from rows of records, row0 is boolean, rowid0 is ID, rowcnt is count
        '        Public Const AdminActionSaveAddNew = 31         ' Save Record and add a new record
        '        Public Const AdminActionReloadCDef = 32         ' Load Content Definitions
        '        Public Const AdminActionWorkflowPublishSelected = 33 ' Publish what was selected
        '        Public Const AdminActionMarkReviewed = 34       ' Mark the record reviewed without making any changes
        '        Public Const AdminActionEditRefresh = 35        ' reload the page just like a save, but do not save
        '        '
        '        ' ----- Adminforms (0-99)
        '        '
        '        Public Const AdminFormRoot = 0             ' intro page
        '        Public Const AdminFormIndex = 1            ' record list page
        '        Public Const AdminFormHelp = 2             ' popup help window
        '        Public Const AdminFormUpload = 3           ' encoded file upload form
        '        Public Const AdminFormEdit = 4             ' Edit form for system format records
        '        Public Const AdminFormEditSystem = 5       ' Edit form for system format records
        '        Public Const AdminFormEditNormal = 6       ' record edit page
        '        Public Const AdminFormEditEmail = 7        ' Edit form for Email format records
        '        Public Const AdminFormEditMember = 8       ' Edit form for Member format records
        '        Public Const AdminFormEditPaths = 9        ' Edit form for Paths format records
        '        Public Const AdminFormClose = 10           ' Special Case - do a window close instead of displaying a form
        '        Public Const AdminFormReports = 12         ' Call Reports form (admin only)
        '        'Public Const AdminFormSpider = 13          ' Call Spider
        '        Public Const AdminFormEditContent = 14     ' Edit form for Content records
        '        Public Const AdminFormDHTMLEdit = 15       ' ActiveX DHTMLEdit form
        '        Public Const AdminFormEditPageContent = 16 '
        '        Public Const AdminFormPublishing = 17       ' Workflow Authoring Publish Control form
        '        Public Const AdminFormQuickStats = 18       ' Quick Stats (from Admin root)
        '        Public Const AdminFormResourceLibrary = 19  ' Resource Library without Selects
        '        Public Const AdminFormEDGControl = 20       ' Control Form for the EDG publishing controls
        '        Public Const AdminFormSpiderControl = 21    ' Control Form for the Content Spider
        '        Public Const AdminFormContentChildTool = 22 ' Admin Create Content Child tool
        '        Public Const AdminformPageContentMap = 23   ' Map all content to a single map
        '        Public Const AdminformHousekeepingControl = 24 ' Housekeeping control
        '        Public Const AdminFormCommerceControl = 25
        '        Public Const AdminFormContactManager = 26
        '        Public Const AdminFormStyleEditor = 27
        '        Public Const AdminFormEmailControl = 28
        '        Public Const AdminFormCommerceInterface = 29
        '        Public Const AdminFormDownloads = 30
        '        Public Const AdminformRSSControl = 31
        '        Public Const AdminFormMeetingSmart = 32
        '        Public Const AdminFormMemberSmart = 33
        '        Public Const AdminFormEmailWizard = 34
        '        Public Const AdminFormImportWizard = 35
        '        Public Const AdminFormCustomReports = 36
        '        Public Const AdminFormFormWizard = 37
        '        Public Const AdminFormLegacyAddonManager = 38
        '        Public Const AdminFormIndex_SubFormAdvancedSearch = 39
        '        Public Const AdminFormIndex_SubFormSetColumns = 40
        '        Public Const AdminFormPageControl = 41
        '        Public Const AdminFormSecurityControl = 42
        '        Public Const AdminFormEditorConfig = 43
        '        Public Const AdminFormBuilderCollection = 44
        '        Public Const AdminFormClearCache = 45
        '        Public Const AdminFormMobileBrowserControl = 46
        '        Public Const AdminFormMetaKeywordTool = 47
        '        Public Const AdminFormIndex_SubFormExport = 48
        '        '
        '        ' ----- AdminFormTools (11,100-199)
        '        '
        '        Public Const AdminFormTools = 11           ' Call Tools form (developer only)
        '        Public Const AdminFormToolRoot = 11         ' These should match for compatibility
        '        Public Const AdminFormToolCreateContentDefinition = 101
        '        Public Const AdminFormToolContentTest = 102
        '        Public Const AdminFormToolConfigureMenu = 103
        '        Public Const AdminFormToolConfigureListing = 104
        '        Public Const AdminFormToolConfigureEdit = 105
        '        Public Const AdminFormToolManualQuery = 106
        '        Public Const AdminFormToolWriteUpdateMacro = 107
        '        Public Const AdminFormToolDuplicateContent = 108
        '        Public Const AdminFormToolDuplicateDataSource = 109
        '        Public Const AdminFormToolDefineContentFieldsFromTable = 110
        '        Public Const AdminFormToolContentDiagnostic = 111
        '        Public Const AdminFormToolCreateChildContent = 112
        '        Public Const AdminFormToolClearContentWatchLink = 113
        '        Public Const AdminFormToolSyncTables = 114
        '        Public Const AdminFormToolBenchmark = 115
        '        Public Const AdminFormToolSchema = 116
        '        Public Const AdminFormToolContentFileView = 117
        '        Public Const AdminFormToolDbIndex = 118
        '        Public Const AdminFormToolContentDbSchema = 119
        '        Public Const AdminFormToolLogFileView = 120
        '        Public Const AdminFormToolLoadCDef = 121
        '        Public Const AdminFormToolLoadTemplates = 122
        '        Public Const AdminformToolFindAndReplace = 123
        '        Public Const AdminformToolCreateGUID = 124
        '        Public Const AdminformToolIISReset = 125
        '        Public Const AdminFormToolRestart = 126
        '        Public Const AdminFormToolWebsiteFileView = 127
        '        '
        '        ' ----- Define the index column structure
        '        '       IndexColumnVariant( 0, n ) is the first column on the left
        '        '       IndexColumnVariant( 0, IndexColumnField ) = the index into the fields array
        '        '
        '        Public Const IndexColumnField = 0          ' The field displayed in the column
        '        Public Const IndexColumnWIDTH = 1          ' The width of the column
        '        Public Const IndexColumnSORTPRIORITY = 2       ' lowest columns sorts first
        '        Public Const IndexColumnSORTDIRECTION = 3      ' direction of the sort on this column
        '        Public Const IndexColumnSATTRIBUTEMAX = 3      ' the number of attributes here
        '        Public Const IndexColumnsMax = 50
        '        '
        '        ' ----- ReportID Constants, moved to ccCommonModule
        '        '
        '        Public Const ReportFormRoot = 1
        '        Public Const ReportFormDailyVisits = 2
        '        Public Const ReportFormWeeklyVisits = 12
        '        Public Const ReportFormSitePath = 4
        '        Public Const ReportFormSearchKeywords = 5
        '        Public Const ReportFormReferers = 6
        '        Public Const ReportFormBrowserList = 8
        '        Public Const ReportFormAddressList = 9
        '        Public Const ReportFormContentProperties = 14
        '        Public Const ReportFormSurveyList = 15
        '        Public Const ReportFormOrdersList = 13
        '        Public Const ReportFormOrderDetails = 21
        '        Public Const ReportFormVisitorList = 11
        '        Public Const ReportFormMemberDetails = 16
        '        Public Const ReportFormPageList = 10
        '        Public Const ReportFormVisitList = 3
        '        Public Const ReportFormVisitDetails = 17
        '        Public Const ReportFormVisitorDetails = 20
        '        Public Const ReportFormSpiderDocList = 22
        '        Public Const ReportFormSpiderErrorList = 23
        '        Public Const ReportFormEDGDocErrors = 24
        '        Public Const ReportFormDownloadLog = 25
        '        Public Const ReportFormSpiderDocDetails = 26
        '        Public Const ReportFormSurveyDetails = 27
        '        Public Const ReportFormEmailDropList = 28
        '        Public Const ReportFormPageTraffic = 29
        '        Public Const ReportFormPagePerformance = 30
        '        Public Const ReportFormEmailDropDetails = 31
        '        Public Const ReportFormEmailOpenDetails = 32
        '        Public Const ReportFormEmailClickDetails = 33
        '        Public Const ReportFormGroupList = 34
        '        Public Const ReportFormGroupMemberList = 35
        '        Public Const ReportFormTrapList = 36
        '        Public Const ReportFormCount = 36
        '        '
        '        '=============================================================================
        '        ' Page Scope Meetings Related Storage
        '        '=============================================================================
        '        '
        '        Public Const MeetingFormIndex = 0
        '        Public Const MeetingFormAttendees = 1
        '        Public Const MeetingFormLinks = 2
        '        Public Const MeetingFormFacility = 3
        '        Public Const MeetingFormHotel = 4
        '        Public Const MeetingFormDetails = 5
        '        '
        '        '------------------------------------------------------------------------------
        '        ' Form actions
        '        '------------------------------------------------------------------------------
        '        '
        '        ' ----- DataSource Types
        '        '
        '        Public Const DataSourceTypeODBCSQL99 = 0
        '        Public Const DataSourceTypeODBCAccess = 1
        '        Public Const DataSourceTypeODBCSQLServer = 2
        '        Public Const DataSourceTypeODBCMySQL = 3
        '        Public Const DataSourceTypeXMLFile = 4      ' Use MSXML Interface to open a file
        '        '
        '        '------------------------------------------------------------------------------
        '        '   Application Status
        '        '------------------------------------------------------------------------------
        '        '
        '        Public Const ApplicationStatusNotFound = 0
        '        Public Const ApplicationStatusLoadedNotRunning = 1
        '        Public Const ApplicationStatusRunning = 2
        '        Public Const ApplicationStatusStarting = 3
        '        Public Const ApplicationStatusUpgrading = 4
        '        ' Public Const ApplicationStatusConnectionBusy = 5    ' can not open connection because already open
        '        Public Const ApplicationStatusKernelFailure = 6     ' can not create Kernel
        '        Public Const ApplicationStatusNoHostService = 7     ' host service process ID not set
        '        Public Const ApplicationStatusLicenseFailure = 8    ' failed to start because of License failure
        '        Public Const ApplicationStatusDbFailure = 9         ' failed to start because ccSetup table not found
        '        Public Const ApplicationStatusUnknownFailure = 10   ' failed to start because of unknown error, see trace log
        '        Public Const ApplicationStatusDbBad = 11            ' ccContent,ccFields no records found
        '        Public Const ApplicationStatusConnectionObjectFailure = 12 ' Connection Object FAiled
        '        Public Const ApplicationStatusConnectionStringFailure = 13 ' Connection String FAiled to open the ODBC connection
        '        Public Const ApplicationStatusDataSourceFailure = 14 ' DataSource failed to open
        '        Public Const ApplicationStatusDuplicateDomains = 15 ' Can not locate application because there are 1+ apps that match the domain
        '        Public Const ApplicationStatusPaused = 16           ' Running, but all activity is blocked (for backup)
        '        '
        '        ' Document (HTML, graphic, etc) retrieved from site
        '        '
        '        Public Const ResponseHeaderCountMax = 20
        '        Public Const ResponseCookieCountMax = 20
        '        '
        '        ' ----- text delimiter that divides the text and html parts of an email message stored in the queue folder
        '        '
        '        Public Const EmailTextHTMLDelimiter = vbCrLf & " ----- End Text Begin HTML -----" & vbCrLf
        '        '
        '        '------------------------------------------------------------------------
        '        '   Common RequestName Variables
        '        '------------------------------------------------------------------------
        '        '
        '        Public Const RequestNameDynamicFormID = "dformid"
        '        '
        '        Public Const RequestNameRunAddon = "addonid"
        '        Public Const RequestNameEditReferer = "EditReferer"
        '        Public Const RequestNameRefreshBlock = "ccFormRefreshBlockSN"
        '        Public Const RequestNameCatalogOrder = "CatalogOrderID"
        '        Public Const RequestNameCatalogCategoryID = "CatalogCatID"
        '        Public Const RequestNameCatalogForm = "CatalogFormID"
        '        Public Const RequestNameCatalogItemID = "CatalogItemID"
        '        Public Const RequestNameCatalogItemAge = "CatalogItemAge"
        '        Public Const RequestNameCatalogRecordTop = "CatalogTop"
        '        Public Const RequestNameCatalogFeatured = "CatalogFeatured"
        '        Public Const RequestNameCatalogSpan = "CatalogSpan"
        '        Public Const RequestNameCatalogKeywords = "CatalogKeywords"
        '        Public Const RequestNameCatalogSource = "CatalogSource"
        '        '
        '        Public Const RequestNameLibraryFileID = "fileEID"
        '        Public Const RequestNameDownloadID = "downloadid"
        '        Public Const RequestNameLibraryUpload = "LibraryUpload"
        '        Public Const RequestNameLibraryName = "LibraryName"
        '        Public Const RequestNameLibraryDescription = "LibraryDescription"

        '        Public Const RequestNameTestHook = "CC"       ' input request that sets debugging hooks

        '        Public Const RequestNameRootPage = "RootPageName"
        '        Public Const RequestNameRootPageID = "RootPageID"
        '        Public Const RequestNameContent = "ContentName"
        '        Public Const RequestNameOrderByClause = "OrderByClause"
        '        Public Const RequestNameAllowChildPageList = "AllowChildPageList"
        '        '
        '        Public Const RequestNameCRKey = "crkey"
        '        Public Const RequestNameAdminForm = "af"
        '        Public Const RequestNameAdminSubForm = "subform"
        '        Public Const RequestNameButton = "button"
        '        Public Const RequestNameAdminSourceForm = "asf"
        '        Public Const RequestNameAdminFormSpelling = "SpellingRequest"
        '        Public Const RequestNameInlineStyles = "InlineStyles"
        '        Public Const RequestNameAllowCSSReset = "AllowCSSReset"
        '        '
        '        Public Const RequestNameReportForm = "rid"
        '        '
        '        Public Const RequestNameToolContentID = "ContentID"
        '        '
        '        Public Const RequestNameCut = "a904o2pa0cut"
        '        Public Const RequestNamePaste = "dp29a7dsa6paste"
        '        Public Const RequestNamePasteParentContentID = "dp29a7dsa6cid"
        '        Public Const RequestNamePasteParentRecordID = "dp29a7dsa6rid"
        '        Public Const RequestNamePasteFieldList = "dp29a7dsa6key"
        '        Public Const RequestNameCutClear = "dp29a7dsa6clear"
        '        '
        '        Public Const RequestNameRequestBinary = "RequestBinary"
        '        ' removed -- this was an old method of blocking form input for file uploads
        '        'Public Const RequestNameFormBlock = "RB"
        '        Public Const RequestNameJSForm = "RequestJSForm"
        '        Public Const RequestNameJSProcess = "ProcessJSForm"
        '        '
        '        Public Const RequestNameFolderID = "FolderID"
        '        '
        '        Public Const RequestNameEmailMemberID = "emi8s9Kj"
        '        Public Const RequestNameEmailOpenFlag = "eof9as88"
        '        Public Const RequestNameEmailOpenCssFlag = "8aa41pM3"
        '        Public Const RequestNameEmailClickFlag = "ecf34Msi"
        '        Public Const RequestNameEmailSpamFlag = "9dq8Nh61"
        '        Public Const RequestNameEmailBlockRequestDropID = "BlockEmailRequest"
        '        Public Const RequestNameVisitTracking = "s9lD1088"
        '        Public Const RequestNameBlockContentTracking = "BlockContentTracking"
        '        Public Const RequestNameCookieDetectVisitID = "f92vo2a8d"

        '        Public Const RequestNamePageNumber = "PageNumber"
        '        Public Const RequestNamePageSize = "PageSize"
        '        '
        '        Public Const RequestValueNull = "[NULL]"
        '        '
        '        Public Const SpellCheckUserDictionaryFilename = "SpellCheck\UserDictionary.txt"
        '        '
        '        Public Const RequestNameStateString = "vstate"
        '        '
        '        '------------------------------------------------------------------------------
        '        ' name value pairs
        '        '------------------------------------------------------------------------------
        '        '
        'Public Type NameValuePairType
        '    Name As String
        '    Value As String
        '    End Type
        '        ''
        '        '' ----- ContentSetMirror Type
        '        ''       Used on the WebClient, not the CSv
        '        ''       Stores info about the ContentSet, and caches the current row
        '        ''
        '        'Public Type ContentSetMirrorType
        '        '    Open As Boolean                     ' If true, it is in use
        '        '    Updateable As Boolean               ' Can not update an OpenCSSQL because Fields are not accessable
        '        '    ContentName As String               ' If updateable, this is the contentname
        '        '    CSPointer as integer                ' CSPointer for this ContentSet
        '        '    '
        '        '    ' ----- a cache of the current row, passed in during open and nextrecord, back during save and nextrecord
        '        '    '
        '        '    EOF As Boolean                      ' if true, Row is empty and at end of records
        '        '    RowCache() As ContentSetRowCacheType ' array of fields buffered for this set
        '        '    RowCacheSize as integer             ' the total number of fields in the row
        '        '    RowCacheCount as integer            ' the number of field() values to write
        '        '    End Type
        '        '
        '        ' ----- Dataset for graphing
        '        '
        'Public Type ColumnDataType
        '    Name As String
        '    row() as integer
        '    End Type
        '        '
        'Public Type ChartDataType
        '    Title As String
        '    XLabel As String
        '    YLabel As String
        '    RowCount as integer
        '    RowLabel() As String
        '    ColumnCount as integer
        '    Column() As ColumnDataType
        '    End Type
        '        ''
        '        ' PrivateStorage to hold the DebugTimer
        '        '
        'Type TimerStackType
        '    Label As String
        '    StartTicks as integer
        '    End Type
        '        Private Const TimerStackMax = 20
        '        Private TimerStack(TimerStackMax) As TimerStackType
        '        Private TimerStackCount as integer
        '        '
        '        Public Const TextSearchStartTagDefault = "<!--TextSearchStart-->"
        '        Public Const TextSearchEndTagDefault = "<!--TextSearchEnd-->"
        '        '
        '        '-------------------------------------------------------------------------------------
        '        '   IPDaemon communication objects
        '        '-------------------------------------------------------------------------------------
        '        '
        'Type IPDaemonConnectionType
        '    ConnectionID As Integer
        '    BytesToSend as integer
        '    HTTPVersion As String
        '    HTTPMethod As String
        '    Path As String
        '    Query As String
        '    Headers As String
        '    PostData As String
        '    SendData As Boolean
        '    State As Integer
        '    ContentLength As Integer
        '    End Type

        'Global IPDaemonConnection() As IPDaemonConnectionType

        'Global Const IPDaemon_DISCONNECTED = 0
        'Global Const IPDaemon_CONNECTED = 1
        'Global Const IPDaemon_HEADERS = 2
        'Global Const IPDaemon_POSTDATA = 3
        'Global Const IPDaemon_SERVE = 4
        'Global Const IPDaemon_SERVEDIR = 5
        'Global Const IPDaemon_SERVEFILE = 6
        '        '
        '        '-------------------------------------------------------------------------------------
        '        '   Email
        '        '-------------------------------------------------------------------------------------
        '        '
        '        Public Const EmailLogTypeDrop = 1                   ' Email was dropped
        '        Public Const EmailLogTypeOpen = 2                   ' System detected the email was opened
        '        Public Const EmailLogTypeClick = 3                  ' System detected a click from a link on the email
        '        Public Const EmailLogTypeBounce = 4                 ' Email was processed by bounce processing
        '        Public Const EmailLogTypeBlockRequest = 5           ' recipient asked us to stop sending email
        '        Public Const EmailLogTypeImmediateSend = 6        ' Email was dropped
        '        '
        '        Public Const DefaultSpamFooter = "<p>To block future emails from this site, <link>click here</link></p>"
        '        '
        '        Public Const FeedbackFormNotSupportedComment = "<!--" & vbCrLf & "Feedback form is not supported in this context" & vbCrLf & "-->"
        '        '
        '        '-------------------------------------------------------------------------------------
        '        '   Page Content constants
        '        '-------------------------------------------------------------------------------------
        '        '
        '        Public Const ContentBlockCopyName = "Content Block Copy"
        '        '
        '        Public Const BubbleCopy_AdminAddPage = "Use the Add page to create new content records. The save button puts you in edit mode. The OK button creates the record and exits."
        '        Public Const BubbleCopy_AdminIndexPage = "Use the Admin Listing page to locate content records through the Admin Site."
        '        Public Const BubbleCopy_SpellCheckPage = "Use the Spell Check page to verify and correct spelling throught the content."
        '        Public Const BubbleCopy_AdminEditPage = "Use the Edit page to add and modify content."
        '        '
        '        '
        '        Public Const TemplateDefaultName = "Default"
        '        'Public Const TemplateDefaultBody = "<!--" & vbCrLf & "Default Template - edit this Page Template, or select a different template for your page or section" & vbCrLf & "-->{{DYNAMICMENU?MENU=}}<br>{{CONTENT}}"
        '        Public Const TemplateDefaultBody = "" _
        '            & vbCrLf & vbTab & "<!--" _
        '            & vbCrLf & vbTab & "Default Template - edit this Page Template, or select a different template for your page or section" _
        '            & vbCrLf & vbTab & "-->" _
        '            & vbCrLf & vbTab & "<ac type=""AGGREGATEFUNCTION"" name=""Dynamic Menu"" querystring=""Menu Name=Default"" acinstanceid=""{6CBADABB-5B0D-43E1-B3CA-46A3D60DA3E1}"" >" _
        '            & vbCrLf & vbTab & "<ac type=""AGGREGATEFUNCTION"" name=""Content Box"" acinstanceid=""{49E0D0C0-D323-49B6-B211-B9599673A265}"" >"
        '        Public Const TemplateDefaultBodyTag = "<body class=""ccBodyWeb"">"
        '        '
        '        '=======================================================================
        '        '   Internal Tab interface storage
        '        '=======================================================================
        '        '
        'Private Type TabType
        '    Caption As String
        '    Link As String
        '    StylePrefix As String
        '    IsHit As Boolean
        '    LiveBody As String
        'End Type
        '        Private Tabs() As TabType
        '        Private TabsCnt as integer
        '        Private TabsSize as integer
        '        '
        '        ' Admin Navigator Nodes
        '        '
        '        Public Const NavigatorNodeCollectionList = -1
        '        Public Const NavigatorNodeAddonList = -1
        '        '
        '        ' Pointers into index of PCC (Page Content Cache) array
        '        '
        '        Public Const PCC_ID = 0
        '        Public Const PCC_Active = 1
        '        Public Const PCC_ParentID = 2
        '        Public Const PCC_Name = 3
        '        Public Const PCC_Headline = 4
        '        Public Const PCC_MenuHeadline = 5
        '        Public Const PCC_DateArchive = 6
        '        Public Const PCC_DateExpires = 7
        '        Public Const PCC_PubDate = 8
        '        Public Const PCC_ChildListSortMethodID = 9
        '        Public Const PCC_ContentControlID = 10
        '        Public Const PCC_TemplateID = 11
        '        Public Const PCC_BlockContent = 12
        '        Public Const PCC_BlockPage = 13
        '        Public Const PCC_Link = 14
        '        Public Const PCC_RegistrationGroupID = 15
        '        Public Const PCC_BlockSourceID = 16
        '        Public Const PCC_CustomBlockMessageFilename = 17
        '        Public Const PCC_JSOnLoad = 18
        '        Public Const PCC_JSHead = 19
        '        Public Const PCC_JSEndBody = 20
        '        Public Const PCC_Viewings = 21
        '        Public Const PCC_ContactMemberID = 22
        '        Public Const PCC_AllowHitNotification = 23
        '        Public Const PCC_TriggerSendSystemEmailID = 24
        '        Public Const PCC_TriggerConditionID = 25
        '        Public Const PCC_TriggerConditionGroupID = 26
        '        Public Const PCC_TriggerAddGroupID = 27
        '        Public Const PCC_TriggerRemoveGroupID = 28
        '        Public Const PCC_AllowMetaContentNoFollow = 29
        '        Public Const PCC_ParentListName = 30
        '        Public Const PCC_CopyFilename = 31
        '        Public Const PCC_BriefFilename = 32
        '        Public Const PCC_AllowChildListDisplay = 33
        '        Public Const PCC_SortOrder = 34
        '        Public Const PCC_DateAdded = 35
        '        Public Const PCC_ModifiedDate = 36
        '        Public Const PCC_ChildPagesFound = 37
        '        Public Const PCC_AllowInMenus = 38
        '        Public Const PCC_AllowInChildLists = 39
        '        Public Const PCC_JSFilename = 40
        '        Public Const PCC_ChildListInstanceOptions = 41
        '        Public Const PCC_IsSecure = 42
        '        Public Const PCC_AllowBrief = 43
        '        Public Const PCC_ColCnt = 44
        '        '
        '        ' Indexes into the SiteSectionCache
        '        ' Created from "ID, Name,TemplateID,ContentID,MenuImageFilename,Caption,MenuImageOverFilename,HideMenu,BlockSection,RootPageID,JSOnLoad,JSHead,JSEndBody"
        '        '
        '        Public Const SSC_ID = 0
        '        Public Const SSC_Name = 1
        '        Public Const SSC_TemplateID = 2
        '        Public Const SSC_ContentID = 3
        '        Public Const SSC_MenuImageFilename = 4
        '        Public Const SSC_Caption = 5
        '        Public Const SSC_MenuImageOverFilename = 6
        '        Public Const SSC_HideMenu = 7
        '        Public Const SSC_BlockSection = 8
        '        Public Const SSC_RootPageID = 9
        '        Public Const SSC_JSOnLoad = 10
        '        Public Const SSC_JSHead = 11
        '        Public Const SSC_JSEndBody = 12
        '        Public Const SSC_JSFilename = 13
        '        Public Const SSC_cnt = 14
        '        '
        '        ' Indexes into the TemplateCache
        '        ' Created from "t.ID,t.Name,t.Link,t.BodyHTML,t.JSOnLoad,t.JSHead,t.JSEndBody,t.StylesFilename,r.StyleID"
        '        '
        '        Public Const TC_ID = 0
        '        Public Const TC_Name = 1
        '        Public Const TC_Link = 2
        '        Public Const TC_BodyHTML = 3
        '        Public Const TC_JSOnLoad = 4
        '        Public Const TC_JSInHeadLegacy = 5
        '        'Public Const TC_JSHead = 5
        '        Public Const TC_JSEndBody = 6
        '        Public Const TC_StylesFilename = 7
        '        Public Const TC_SharedStylesIDList = 8
        '        Public Const TC_MobileBodyHTML = 9
        '        Public Const TC_MobileStylesFilename = 10
        '        Public Const TC_OtherHeadTags = 11
        '        Public Const TC_BodyTag = 12
        '        Public Const TC_JSInHeadFilename = 13
        '        'Public Const TC_JSFilename = 13
        '        Public Const TC_IsSecure = 14
        '        Public Const TC_DomainIdList = 15
        '        ' for now, Mobile templates do not have shared styles
        '        'Public Const TC_MobileSharedStylesIDList = 11
        '        Public Const TC_cnt = 16
        '        '
        '        ' DTD
        '        '
        '        Public Const DTDDefault = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">"
        '        Public Const DTDDefaultAdmin = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">"
        '        '
        '        ' innova Editor feature list
        '        '
        '        Public Const InnovaEditorFeaturefilename = "Config\EditorCongif.txt"
        '        Public Const InnovaEditorFeatureList = "FullScreen,Preview,Print,Search,Cut,Copy,Paste,PasteWord,PasteText,SpellCheck,Undo,Redo,Image,Flash,Media,CustomObject,CustomTag,Bookmark,Hyperlink,HTMLSource,XHTMLSource,Numbering,Bullets,Indent,Outdent,JustifyLeft,JustifyCenter,JustifyRight,JustifyFull,Table,Guidelines,Absolute,Characters,Line,Form,RemoveFormat,ClearAll,StyleAndFormatting,TextFormatting,ListFormatting,BoxFormatting,ParagraphFormatting,CssText,Styles,Paragraph,FontName,FontSize,Bold,Italic,Underline,Strikethrough,Superscript,Subscript,ForeColor,BackColor"
        '        Public Const InnovaEditorPublicFeatureList = "FullScreen,Preview,Print,Search,Cut,Copy,Paste,PasteWord,PasteText,SpellCheck,Undo,Redo,Bookmark,Hyperlink,HTMLSource,XHTMLSource,Numbering,Bullets,Indent,Outdent,JustifyLeft,JustifyCenter,JustifyRight,JustifyFull,Table,Guidelines,Absolute,Characters,Line,Form,RemoveFormat,ClearAll,StyleAndFormatting,TextFormatting,ListFormatting,BoxFormatting,ParagraphFormatting,CssText,Styles,Paragraph,FontName,FontSize,Bold,Italic,Underline,Strikethrough,Superscript,Subscript,ForeColor,BackColor"
        '        ''
        '        '' Content Type
        '        ''
        '        'Enum contentTypeEnum
        '        '    contentTypeWeb = 1
        '        '    ContentTypeEmail = 2
        '        '    contentTypeWebTemplate = 3
        '        '    contentTypeEmailTemplate = 4
        '        'End Enum
        '        'Public EditorContext As contentTypeEnum
        '        'Enum EditorContextEnum
        '        '    contentTypeWeb = 1
        '        '    contentTypeEmail = 2
        '        'End Enum
        '        'Public EditorContext As EditorContextEnum
        '        ''
        '        'Public Const EditorAddonMenuEmailTemplateFilename = "templates/EditorAddonMenuTemplateEmail.js"
        '        'Public Const EditorAddonMenuEmailContentFilename = "templates/EditorAddonMenuContentEmail.js"
        '        'Public Const EditorAddonMenuWebTemplateFilename = "templates/EditorAddonMenuTemplateWeb.js"
        '        'Public Const EditorAddonMenuWebContentFilename = "templates/EditorAddonMenuContentWeb.js"
        '        '
        '        Public Const DynamicStylesFilename = "templates/styles.css"
        '        Public Const AdminSiteStylesFilename = "templates/AdminSiteStyles.css"
        '        Public Const EditorStyleRulesFilenamePattern = "templates/EditorStyleRules$TemplateID$.js"
        '        ' deprecated 11/24/3009 - StyleRules destinction between web/email not needed b/c body background blocked
        '        'Public Const EditorStyleWebRulesFilename = "templates/EditorStyleWebRules.js"
        '        'Public Const EditorStyleEmailRulesFilename = "templates/EditorStyleEmailRules.js"
        '        '
        '        ' ----- ccGroupRules storage for list of Content that a group can author
        '        '
        'Public Type ContentGroupRuleType
        '    ContentID as integer
        '    GroupID as integer
        '    AllowAdd As Boolean
        '    AllowDelete As Boolean
        'End Type
        '        '
        '        ' ----- This should match the Lookup List in the NavIconType field in the Navigator Entry content definition
        '        '
        '        Public Const navTypeIDList = "Add-on,Report,Setting,Tool"
        '        Public Const NavTypeIDAddon = 1
        '        Public Const NavTypeIDReport = 2
        '        Public Const NavTypeIDSetting = 3
        '        Public Const NavTypeIDTool = 4
        '        '
        Public Const NavIconTypeList = "Custom,Advanced,Content,Folder,Email,User,Report,Setting,Tool,Record,Addon,help"
        Public Const NavIconTypeCustom As Integer = 1
        Public Const NavIconTypeAdvanced As Integer = 2
        Public Const NavIconTypeContent As Integer = 3
        Public Const NavIconTypeFolder As Integer = 4
        Public Const NavIconTypeEmail As Integer = 5
        Public Const NavIconTypeUser As Integer = 6
        Public Const NavIconTypeReport As Integer = 7
        Public Const NavIconTypeSetting As Integer = 8
        Public Const NavIconTypeTool As Integer = 9
        Public Const NavIconTypeRecord As Integer = 10
        Public Const NavIconTypeAddon As Integer = 11
        Public Const NavIconTypeHelp As Integer = 12
        '        '
        '        Public Const QueryTypeSQL = 1
        '        Public Const QueryTypeOpenContent = 2
        '        Public Const QueryTypeUpdateContent = 3
        '        Public Const QueryTypeInsertContent = 4
        '        '
        '        ' Google Data Object construction in GetRemoteQuery
        '        '
        'Public Type ColsType
        '    Type As String
        '    Id As String
        '    Label As String
        '    Pattern As String
        'End Type
        '        '
        'Public Type CellType
        '    v As String
        '    f As String
        '    p As String
        'End Type
        '        '
        'Public Type RowsType
        '    Cell() As CellType
        'End Type
        '        '
        'Public Type GoogleDataType
        '    IsEmpty As Boolean
        '    col() As ColsType
        '    row() As RowsType
        'End Type
        '        '
        '        Public Enum GoogleVisualizationStatusEnum
        '            OK = 1
        '            warning = 2
        '    Error = 3
        'End Enum
        '        '
        'Public Type GoogleVisualizationType
        '    version As String
        '    reqid As String
        '    status As GoogleVisualizationStatusEnum
        '    warnings() As String
        '    errors() As String
        '    sig As String
        '    table As GoogleDataType
        'End Type

        '        'Public Const ReturnFormatTypeGoogleTable = 1
        '        'Public Const ReturnFormatTypeNameValue = 2

        '        Public Enum RemoteFormatEnum
        '            RemoteFormatJsonTable = 1
        '            RemoteFormatJsonNameArray = 2
        '            RemoteFormatJsonNameValue = 3
        '        End Enum
        '        '
        '        '
        '        '
        '        Public Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
        '        Public Declare Function RegOpenKeyExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpszSubKey$, dwOptions&, ByVal samDesired&, lpHKey&)
        '        Public Declare Function RegQueryValueExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)
        '        Public Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, lpDataBuff&, nSize&)

        '        Public Const HKEY_CLASSES_ROOT = &H80000000
        '        Public Const HKEY_CURRENT_USER = &H80000001
        '        Public Const HKEY_LOCAL_MACHINE = &H80000002
        '        Public Const HKEY_USERS = &H80000003

        '        Public Const ERROR_SUCCESS = 0&
        '        Public Const REG_SZ = 1&                          ' Unicode nul terminated string
        '        Public Const REG_DWORD = 4&                       ' 32-bit number

        '        Public Const KEY_QUERY_VALUE = &H1&
        '        Public Const KEY_SET_VALUE = &H2&
        '        Public Const KEY_CREATE_SUB_KEY = &H4&
        '        Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
        '        Public Const KEY_NOTIFY = &H10&
        '        Public Const KEY_CREATE_LINK = &H20&
        '        Public Const READ_CONTROL = &H20000
        '        Public Const WRITE_DAC = &H40000
        '        Public Const WRITE_OWNER = &H80000
        '        Public Const SYNCHRONIZE = &H100000
        '        Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
        '        Public Const STANDARD_RIGHTS_READ = READ_CONTROL
        '        Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
        '        Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
        '        Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
        '        Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
        '        Public Const KEY_EXECUTE = KEY_READ

        '        '======================================================================================
        '        '
        '        '======================================================================================
        '        '
        '        Public Sub StartDebugTimer(Enabled As Boolean, Label As String)
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            If Enabled Then
        '                If TimerStackCount < TimerStackMax Then
        '                    TimerStack(TimerStackCount).Label = Label
        '                    TimerStack(TimerStackCount).StartTicks = GetTickCount
        '                Else
        '                    Call AppendLogFile(App.EXEName & ".?.StartDebugTimer, " & "Timer Stack overflow, attempting push # [" & TimerStackCount & "], but max = [" & TimerStackMax & "]")
        '                End If
        '                TimerStackCount = TimerStackCount + 1
        '            End If
        '        End Sub
        '        '
        '        Public Sub StopDebugTimer(Enabled As Boolean, Label As String)
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            If Enabled Then
        '                If TimerStackCount <= 0 Then
        '                    Call AppendLogFile(App.EXEName & ".?.StopDebugTimer, " & "Timer Error, attempting to Pop, but the stack is empty")
        '                Else
        '                    If TimerStackCount <= TimerStackMax Then
        '                        If TimerStack(TimerStackCount - 1).Label = Label Then
        '                    Call AppendLogFile(App.EXEName & ".?.StopDebugTimer, " & "Timer [" & String(2 * TimerStackCount, ".") & Label & "] took " & (GetTickCount - TimerStack(TimerStackCount - 1).StartTicks) & " msec")
        '                        Else
        '                            Call AppendLogFile(App.EXEName & ".?.StopDebugTimer, " & "Timer Error, [" & Label & "] was popped, but [" & TimerStack(TimerStackCount).Label & "] was on the top of the stack")
        '                        End If
        '                    End If
        '                    TimerStackCount = TimerStackCount - 1
        '                End If
        '            End If
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Function PayString(Index) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            Select Case Index
        '                Case PayTypeCreditCardOnline
        '                    PayString = "Credit Card"
        '                Case PayTypeCreditCardByPhone
        '                    PayString = "Credit Card by phone"
        '                Case PayTypeCreditCardByFax
        '                    PayString = "Credit Card by fax"
        '                Case PayTypeCHECK
        '                    PayString = "Personal Check"
        '                Case PayTypeCHECKCOMPANY
        '                    PayString = "Company Check"
        '                Case PayTypeCREDIT
        '                    PayString = "You will be billed"
        '                Case PayTypeNetTerms
        '                    PayString = "Net Terms (Approved customers only)"
        '                Case PayTypeCODCompanyCheck
        '                    PayString = "COD- Pre-Approved Only"
        '                Case PayTypeCODCertifiedFunds
        '                    PayString = "COD- Certified Funds"
        '                Case PayTypePAYPAL
        '                    PayString = "PayPal"
        '                Case Else
        '                    ' Case PayTypeNONE
        '                    PayString = "No payment required"
        '            End Select
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function CCString(Index) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            Select Case Index
        '                Case CCTYPEVISA
        '                    CCString = "Visa"
        '                Case CCTYPEMC
        '                    CCString = "MasterCard"
        '                Case CCTYPEAMEX
        '                    CCString = "American Express"
        '                Case CCTYPEDISCOVER
        '                    CCString = "Discover"
        '                Case Else
        '                    ' Case CCTYPENOVUS
        '                    CCString = "Novus Card"
        '            End Select
        '        End Function
        '        '
        '        '========================================================================
        '        ' Get a Long from a CommandPacket
        '        '   position+0, 4 byte value
        '        '========================================================================
        '        '
        '        Public Function GetLongFromByteArray(ByteArray() As Byte, Position as integer) as integer
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            GetLongFromByteArray = ByteArray(Position + 3)
        '            GetLongFromByteArray = ByteArray(Position + 2) + (256 * GetLongFromByteArray)
        '            GetLongFromByteArray = ByteArray(Position + 1) + (256 * GetLongFromByteArray)
        '            GetLongFromByteArray = ByteArray(Position + 0) + (256 * GetLongFromByteArray)
        '            Position = Position + 4
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' Get a Long from a byte array
        '        '   position+0, 4 byte size of the number
        '        '   position+3, start of the number
        '        '========================================================================
        '        '
        '        Public Function GetNumberFromByteArray(ByteArray() As Byte, Position as integer) as integer
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim ArgumentCount as integer
        '            Dim ArgumentLength as integer
        '            '
        '            ArgumentLength = GetLongFromByteArray(ByteArray(), Position)
        '            '
        '            If ArgumentLength > 0 Then
        '                GetNumberFromByteArray = 0
        '                For ArgumentCount = ArgumentLength - 1 To 0 Step -1
        '                    GetNumberFromByteArray = ByteArray(Position + ArgumentCount) + (256 * GetNumberFromByteArray)
        '                Next
        '            End If
        '            Position = Position + ArgumentLength
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' Get a String a byte array
        '        '   position+0, 4 byte length of the string
        '        '   position+3, start of the string
        '        '========================================================================
        '        '
        '        Public Function GetStringFromByteArray(ByteArray() As Byte, Position as integer) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim Pointer as integer
        '            Dim ArgumentLength as integer
        '            '
        '            ArgumentLength = GetLongFromByteArray(ByteArray(), Position)
        '            '
        '            GetStringFromByteArray = ""
        '            If ArgumentLength > 0 Then
        '                For Pointer = 0 To ArgumentLength - 1
        '                    GetStringFromByteArray = GetStringFromByteArray & Chr(ByteArray(Position + Pointer))
        '                Next
        '            End If
        '            Position = Position + ArgumentLength
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' Get a Long from a byte array
        '        '========================================================================
        '        '
        '        Public Sub SetLongByteArray(ByRef ByteArray() As Byte, Position as integer, LongValue as integer)
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            ByteArray(Position + 0) = LongValue And (&HFF)
        '            ByteArray(Position + 1) = Int(LongValue / 256) And (&HFF)
        '            ByteArray(Position + 2) = Int(LongValue / (256 ^ 2)) And (&HFF)
        '            ByteArray(Position + 3) = Int(LongValue / (256 ^ 3)) And (&HFF)
        '            Position = Position + 4
        '            '
        '        End Sub
        '        '
        '        '========================================================================
        '        ' Set a string in a byte array
        '        '========================================================================
        '        '
        '        Public Sub SetStringByteArray(ByRef ByteArray() As Byte, Position as integer, StringValue As String)
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim Pointer as integer
        '            Dim LenStringValue as integer
        '            '
        '            LenStringValue = Len(StringValue)
        '            If LenStringValue > 0 Then
        '                For Pointer = 0 To LenStringValue - 1
        '                    ByteArray(Position + Pointer) = Asc(Mid(StringValue, Pointer + 1, 1)) And (&HFF)
        '                Next
        '                Position = Position + LenStringValue
        '            End If
        '            '
        '        End Sub

        '        '
        '        '========================================================================
        '        '   Set a Long long on the end of a RMB (Remote Method Block)
        '        '       You determine the position, or it will add it to the end
        '        '========================================================================
        '        '
        'Public Sub SetRMBLong(ByRef ByteArray() As Byte, LongValue as integer, Optional Position)
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim Temp as integer
        '            Dim MyPosition as integer
        '            Dim ByteArraySize as integer
        '            '
        '            ' ----- determine the position
        '            '
        '            If Not IsMissing(Position) Then
        '                MyPosition = Position
        '            Else
        '                '
        '                ' ----- Add it to the end, determine length
        '                '
        '                MyPosition = ByteArray(RMBPositionLength + 3)
        '                MyPosition = ByteArray(RMBPositionLength + 2) + (256 * MyPosition)
        '                MyPosition = ByteArray(RMBPositionLength + 1) + (256 * MyPosition)
        '                MyPosition = ByteArray(RMBPositionLength + 0) + (256 * MyPosition)
        '                '
        '                ' ----- adjust size of array if necessary
        '                '
        '                ByteArraySize = UBound(ByteArray)
        '                If ByteArraySize < (MyPosition + 8) Then
        '                    ReDim Preserve ByteArray(ByteArraySize + 8)
        '                End If
        '            End If
        '            '
        '            ' ----- set the length
        '            '
        '            'ByteArray(MyPosition + 0) = 4
        '            'ByteArray(MyPosition + 1) = 0
        '            'ByteArray(MyPosition + 2) = 0
        '            'ByteArray(MyPosition + 3) = 0
        '            'MyPosition = MyPosition + 4
        '            '
        '            ' ----- set the value
        '            '
        '            ByteArray(MyPosition + 0) = LongValue And (&HFF)
        '            ByteArray(MyPosition + 1) = Int(LongValue / 256) And (&HFF)
        '            ByteArray(MyPosition + 2) = Int(LongValue / (256 ^ 2)) And (&HFF)
        '            ByteArray(MyPosition + 3) = Int(LongValue / (256 ^ 3)) And (&HFF)
        '            MyPosition = MyPosition + 4
        '            '
        '            If IsMissing(Position) Then
        '                '
        '                ' ----- Adjust the RMB length if length not given
        '                '
        '                ByteArray(RMBPositionLength + 0) = MyPosition And (&HFF)
        '                ByteArray(RMBPositionLength + 1) = Int(MyPosition / 256) And (&HFF)
        '                ByteArray(RMBPositionLength + 2) = Int(MyPosition / (256 ^ 2)) And (&HFF)
        '                ByteArray(RMBPositionLength + 3) = Int(MyPosition / (256 ^ 3)) And (&HFF)
        '            End If
        '            '
        '        End Sub
        '        '
        '        '========================================================================
        '        '   Set a Long long on the end of a RMB (Remote Method Block)
        '        '       You determine the position, or it will add it to the end
        '        '========================================================================
        '        '
        'Public Sub SetRMBString(ByRef ByteArray() As Byte, StringValue As String, Optional Position)
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim Temp as integer
        '            Dim MyPosition as integer
        '            Dim ByteArraySize as integer
        '            '
        '            ' ----- determine the position
        '            '
        '            If Not IsMissing(Position) Then
        '                MyPosition = Position
        '            Else
        '                '
        '                ' ----- Add it to the end, determine length
        '                '
        '                MyPosition = ByteArray(RMBPositionLength + 3)
        '                MyPosition = ByteArray(RMBPositionLength + 2) + (256 * MyPosition)
        '                MyPosition = ByteArray(RMBPositionLength + 1) + (256 * MyPosition)
        '                MyPosition = ByteArray(RMBPositionLength + 0) + (256 * MyPosition)
        '                '
        '                ' ----- adjust size of array if necessary
        '                '
        '                ByteArraySize = UBound(ByteArray)
        '                If ByteArraySize < (MyPosition + 8) Then
        '                    ReDim Preserve ByteArray(ByteArraySize + 4 + Len(StringValue))
        '                End If
        '            End If
        '            '
        '            ' ----- set the value
        '            '

        '            '
        '            Dim Pointer as integer
        '            Dim LenStringValue as integer
        '            '
        '            LenStringValue = Len(StringValue)
        '            If LenStringValue > 0 Then
        '                For Pointer = 0 To LenStringValue - 1
        '                    ByteArray(MyPosition + Pointer) = Asc(Mid(StringValue, Pointer + 1, 1)) And (&HFF)
        '                Next
        '                MyPosition = MyPosition + LenStringValue
        '            End If
        '            '
        '            If IsMissing(Position) Then
        '                '
        '                ' ----- Adjust the RMB length if length not given
        '                '
        '                ByteArray(RMBPositionLength + 0) = MyPosition And (&HFF)
        '                ByteArray(RMBPositionLength + 1) = Int(MyPosition / 256) And (&HFF)
        '                ByteArray(RMBPositionLength + 2) = Int(MyPosition / (256 ^ 2)) And (&HFF)
        '                ByteArray(RMBPositionLength + 3) = Int(MyPosition / (256 ^ 3)) And (&HFF)
        '            End If
        '            '
        '        End Sub
        '        '
        '        '========================================================================
        '        '   IsTrue
        '        '       returns true or false depending on the state of the variant input
        '        '========================================================================
        '        '
        '        Function IsTrue(ValueVariant) As Boolean
        '            IsTrue = cp.utils.encodeboolean(ValueVariant)
        '        End Function
        '        '
        '        '========================================================================
        '        ' EncodeXML
        '        '
        '        '========================================================================
        '        '
        '        Function EncodeXML(ValueVariant As Object, fieldType as integer) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim TimeValuething As Single
        '            Dim TimeHours as integer
        '            Dim TimeMinutes as integer
        '            Dim TimeSeconds as integer
        '            '
        '            Select Case fieldType
        '                Case FieldTypeInteger, FieldTypeLookup, FieldTypeRedirect, FieldTypeManyToMany, FieldTypeMemberSelect
        '                    If IsNull(ValueVariant) Then
        '                        EncodeXML = "null"
        '                    ElseIf ValueVariant = "" Then
        '                        EncodeXML = "null"
        '                    ElseIf IsNumeric(ValueVariant) Then
        '                        EncodeXML = Int(ValueVariant)
        '                    Else
        '                        EncodeXML = "null"
        '                    End If
        '                Case FieldTypeBoolean
        '                    If IsNull(ValueVariant) Then
        '                        EncodeXML = "0"
        '                    ElseIf ValueVariant <> False Then
        '                        EncodeXML = "1"
        '                    Else
        '                        EncodeXML = "0"
        '                    End If
        '                Case FieldTypeCurrency
        '                    If IsNull(ValueVariant) Then
        '                        EncodeXML = "null"
        '                    ElseIf ValueVariant = "" Then
        '                        EncodeXML = "null"
        '                    ElseIf IsNumeric(ValueVariant) Then
        '                        EncodeXML = ValueVariant
        '                    Else
        '                        EncodeXML = "null"
        '                    End If
        '                Case FieldTypeFloat
        '                    If IsNull(ValueVariant) Then
        '                        EncodeXML = "null"
        '                    ElseIf ValueVariant = "" Then
        '                        EncodeXML = "null"
        '                    ElseIf IsNumeric(ValueVariant) Then
        '                        EncodeXML = ValueVariant
        '                    Else
        '                        EncodeXML = "null"
        '                    End If
        '                Case FieldTypeDate
        '                    If IsNull(ValueVariant) Then
        '                        EncodeXML = "null"
        '                    ElseIf ValueVariant = "" Then
        '                        EncodeXML = "null"
        '                    ElseIf IsDate(ValueVariant) Then
        '                        'TimeVar = CDate(ValueVariant)
        '                        'TimeValuething = 86400! * (TimeVar - Int(TimeVar))
        '                        'TimeHours = Int(TimeValuething / 3600!)
        '                        'TimeMinutes = Int(TimeValuething / 60!) - (TimeHours * 60)
        '                        'TimeSeconds = TimeValuething - (TimeHours * 3600!) - (TimeMinutes * 60!)
        '                        'EncodeXML = Year(ValueVariant) & "-" & Right("0" & Month(ValueVariant), 2) & "-" & Right("0" & Day(ValueVariant), 2) & " " & Right("0" & TimeHours, 2) & ":" & Right("0" & TimeMinutes, 2) & ":" & Right("0" & TimeSeconds, 2)
        '                        EncodeXML = kmaEncodeText(ValueVariant)
        '                    End If
        '                Case Else
        '                    '
        '                    ' ----- FieldTypeText
        '                    ' ----- FieldTypeLongText
        '                    ' ----- FieldTypeFile
        '                    ' ----- FieldTypeImage
        '                    ' ----- FieldTypeTextFile
        '                    ' ----- FieldTypeCSSFile
        '                    ' ----- FieldTypeXMLFile
        '                    ' ----- FieldTypeJavascriptFile
        '                    ' ----- FieldTypeLink
        '                    ' ----- FieldTypeResourceLink
        '                    ' ----- FieldTypeHTML
        '                    ' ----- FieldTypeHTMLFile
        '                    '
        '                    If IsNull(ValueVariant) Then
        '                        EncodeXML = "null"
        '                    ElseIf ValueVariant = "" Then
        '                        EncodeXML = ""
        '                    Else
        '                        'EncodeXML = ASPServer.HTMLEncode(ValueVariant)
        '                        'EncodeXML = Replace(ValueVariant, "&", "&lt;")
        '                        'EncodeXML = Replace(ValueVariant, "<", "&lt;")
        '                        'EncodeXML = Replace(EncodeXML, ">", "&gt;")
        '                    End If
        '            End Select
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' EncodeFilename
        '        '
        '        '========================================================================
        '        '
        '        Public Function encodeFilename(Source As String) As String
        '            Dim allowed As String
        '            Dim chr As String
        '            Dim Ptr as integer
        '            Dim cnt as integer
        '            Dim returnString As String
        '            '
        '            returnString = ""
        '            cnt = Len(Source)
        '            If cnt > 254 Then
        '                cnt = 254
        '            End If
        '            allowed = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ^&'@{}[],$-#()%.+~_"
        '            For Ptr = 1 To cnt
        '                chr = Mid(Source, Ptr, 1)
        '                If InStr(1, allowed, chr, vbBinaryCompare) Then
        '                    returnString = returnString & chr
        '                End If
        '            Next
        '            encodeFilename = returnString
        '        End Function
        '        '
        '        'Function encodeFilename(Filename As String) As String
        '        '    ' ##### removed to catch err<>0 problem on error resume next
        '        '    '
        '        '    Dim Source() As Variant
        '        '    Dim Replacement() As Variant
        '        '    '
        '        '    Source = Array("""", "*", "/", ":", "<", ">", "?", "\", "|", "=")
        '        '    Replacement = Array("_", "_", "_", "_", "_", "_", "_", "_", "_", "_")
        '        '    '
        '        '    encodeFilename = ReplaceMany(Filename, Source, Replacement)
        '        '    If Len(encodeFilename) > 254 Then
        '        '        encodeFilename = Left(encodeFilename, 254)
        '        '    End If
        '        '    encodeFilename = Replace(encodeFilename, vbCr, "_")
        '        '    encodeFilename = Replace(encodeFilename, vbLf, "_")
        '        '    '
        '        '    End Function
        '        '
        '        '
        '        '

        '        '
        '        '========================================================================
        '        ' DecodeHTML
        '        '
        '        '========================================================================
        '        '
        '        Function DecodeHTML(Source As String) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            DecodeHTML = kmaDecodeHTML(Source)
        '            'Dim SourceChr() As Variant
        '            'Dim ReplacementChr() As Variant
        '            ''
        '            'SourceChr = Array("&gt;", "&lt;", "&nbsp;", "&amp;")
        '            'ReplacementChr = Array(">", "<", " ", "&")
        '            ''
        '            'DecodeHTML = ReplaceMany(Source, SourceChr, ReplacementChr)
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        ' EncodeFilename
        '        '
        '        '========================================================================
        '        '
        '        Function ReplaceMany(Source As String, ArrayOfSource() As Object, ArrayOfReplacement() As Object) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim Count as integer
        '            Dim Pointer as integer
        '            '
        '            Count = UBound(ArrayOfSource) + 1
        '            ReplaceMany = Source
        '            For Pointer = 0 To Count - 1
        '                ReplaceMany = Replace(ReplaceMany, ArrayOfSource(Pointer), ArrayOfReplacement(Pointer))
        '            Next
        '            '
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetURIHost(URI) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            '   Divide the URI into URIHost, URIPath, and URIPage
        '            '
        '            Dim URIWorking As String
        '            Dim Slash as integer
        '            Dim LastSlash as integer
        '            Dim URIHost As String
        '            Dim URIPath As String
        '            Dim URIPage As String
        '            URIWorking = URI
        '            If Mid(UCase(URIWorking), 1, 4) = "HTTP" Then
        '                URIWorking = Mid(URIWorking, InStr(1, URIWorking, "//") + 2)
        '            End If
        '            URIHost = URIWorking
        '            Slash = InStr(1, URIHost, "/")
        '            If Slash = 0 Then
        '                URIPath = "/"
        '                URIPage = ""
        '            Else
        '                URIPath = Mid(URIHost, Slash)
        '                URIHost = Mid(URIHost, 1, Slash - 1)
        '                Slash = InStr(1, URIPath, "/")
        '                Do While Slash <> 0
        '                    LastSlash = Slash
        '                    Slash = InStr(LastSlash + 1, URIPath, "/")
        '                    DoEvents()
        '                Loop
        '                URIPage = Mid(URIPath, LastSlash + 1)
        '                URIPath = Mid(URIPath, 1, LastSlash)
        '            End If
        '            GetURIHost = URIHost
        '            '
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetURIPage(URI) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            '   Divide the URI into URIHost, URIPath, and URIPage
        '            '
        '            Dim Slash as integer
        '            Dim LastSlash as integer
        '            Dim URIHost As String
        '            Dim URIPath As String
        '            Dim URIPage As String
        '            Dim URIWorking As String
        '            URIWorking = URI
        '            If Mid(UCase(URIWorking), 1, 4) = "HTTP" Then
        '                URIWorking = Mid(URIWorking, InStr(1, URIWorking, "//") + 2)
        '            End If
        '            URIHost = URIWorking
        '            Slash = InStr(1, URIHost, "/")
        '            If Slash = 0 Then
        '                URIPath = "/"
        '                URIPage = ""
        '            Else
        '                URIPath = Mid(URIHost, Slash)
        '                URIHost = Mid(URIHost, 1, Slash - 1)
        '                Slash = InStr(1, URIPath, "/")
        '                Do While Slash <> 0
        '                    LastSlash = Slash
        '                    Slash = InStr(LastSlash + 1, URIPath, "/")
        '                    DoEvents()
        '                Loop
        '                URIPage = Mid(URIPath, LastSlash + 1)
        '                URIPath = Mid(URIPath, 1, LastSlash)
        '            End If
        '            GetURIPage = URIPage
        '            '
        '        End Function
        '        '
        '        '
        '        '
        '        Function GetDateFromGMT(GMTDate As String) As Date
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim WorkString As String
        '            GetDateFromGMT = 0
        '            If GMTDate <> "" Then
        '                WorkString = Mid(GMTDate, 6, 11)
        '                If IsDate(WorkString) Then
        '                    GetDateFromGMT = CDate(WorkString)
        '                    WorkString = Mid(GMTDate, 18, 8)
        '                    If IsDate(WorkString) Then
        '                        GetDateFromGMT = GetDateFromGMT + CDate(WorkString) + 4 / 24
        '                    End If
        '                End If
        '            End If
        '            '
        '        End Function
        '        '
        '        ' Wdy, DD-Mon-YYYY HH:MM:SS GMT
        '        '
        '        Function GetGMTFromDate(DateValue As Date) As String
        '            '
        '            Dim WorkString As String
        '            Dim WorkLong as integer
        '            '
        '            If IsDate(DateValue) Then
        '                Select Case Weekday(DateValue)
        '                    Case vbSunday
        '                        GetGMTFromDate = "Sun, "
        '                    Case vbMonday
        '                        GetGMTFromDate = "Mon, "
        '                    Case vbTuesday
        '                        GetGMTFromDate = "Tue, "
        '                    Case vbWednesday
        '                        GetGMTFromDate = "Wed, "
        '                    Case vbThursday
        '                        GetGMTFromDate = "Thu, "
        '                    Case vbFriday
        '                        GetGMTFromDate = "Fri, "
        '                    Case vbSaturday
        '                        GetGMTFromDate = "Sat, "
        '                End Select
        '                '
        '                WorkLong = Day(DateValue)
        '                If WorkLong < 10 Then
        '                    GetGMTFromDate = GetGMTFromDate & "0" & CStr(WorkLong) & " "
        '                Else
        '                    GetGMTFromDate = GetGMTFromDate & CStr(WorkLong) & " "
        '                End If
        '                '
        '                Select Case Month(DateValue)
        '                    Case 1
        '                        GetGMTFromDate = GetGMTFromDate & "Jan "
        '                    Case 2
        '                        GetGMTFromDate = GetGMTFromDate & "Feb "
        '                    Case 3
        '                        GetGMTFromDate = GetGMTFromDate & "Mar "
        '                    Case 4
        '                        GetGMTFromDate = GetGMTFromDate & "Apr "
        '                    Case 5
        '                        GetGMTFromDate = GetGMTFromDate & "May "
        '                    Case 6
        '                        GetGMTFromDate = GetGMTFromDate & "Jun "
        '                    Case 7
        '                        GetGMTFromDate = GetGMTFromDate & "Jul "
        '                    Case 8
        '                        GetGMTFromDate = GetGMTFromDate & "Aug "
        '                    Case 9
        '                        GetGMTFromDate = GetGMTFromDate & "Sep "
        '                    Case 10
        '                        GetGMTFromDate = GetGMTFromDate & "Oct "
        '                    Case 11
        '                        GetGMTFromDate = GetGMTFromDate & "Nov "
        '                    Case 12
        '                        GetGMTFromDate = GetGMTFromDate & "Dec "
        '                End Select
        '                '
        '                GetGMTFromDate = GetGMTFromDate & CStr(Year(DateValue)) & " "
        '                '
        '                WorkLong = Hour(DateValue)
        '                If WorkLong < 10 Then
        '                    GetGMTFromDate = GetGMTFromDate & "0" & CStr(WorkLong) & ":"
        '                Else
        '                    GetGMTFromDate = GetGMTFromDate & CStr(WorkLong) & ":"
        '                End If
        '                '
        '                WorkLong = Minute(DateValue)
        '                If WorkLong < 10 Then
        '                    GetGMTFromDate = GetGMTFromDate & "0" & CStr(WorkLong) & ":"
        '                Else
        '                    GetGMTFromDate = GetGMTFromDate & CStr(WorkLong) & ":"
        '                End If
        '                '
        '                WorkLong = Second(DateValue)
        '                If WorkLong < 10 Then
        '                    GetGMTFromDate = GetGMTFromDate & "0" & CStr(WorkLong)
        '                Else
        '                    GetGMTFromDate = GetGMTFromDate & CStr(WorkLong)
        '                End If
        '                '
        '                GetGMTFromDate = GetGMTFromDate & " GMT"
        '            End If
        '            '
        '        End Function
        '        '
        '        '========================================================================
        '        '   EncodeSQL
        '        '       encode a variable to go in an sql expression
        '        '       NOT supported
        '        '========================================================================
        '        '
        'Public Function EncodeSQL(ExpressionVariant As Variant, Optional fieldType As Variant) As String
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim iFieldType as integer
        '            Dim MethodName As String
        '            '
        '            MethodName = "EncodeSQL"
        '            '
        '            iFieldType = KmaEncodeMissingInteger(fieldType, FieldTypeText)
        '            Select Case iFieldType
        '                Case FieldTypeBoolean
        '                    EncodeSQL = KmaEncodeSQLBoolean(ExpressionVariant)
        '                Case FieldTypeCurrency, FieldTypeAutoIncrement, FieldTypeFloat, FieldTypeInteger, FieldTypeLookup, FieldTypeMemberSelect
        '                    EncodeSQL = KmaEncodeSQLNumber(ExpressionVariant)
        '                Case FieldTypeDate
        '                    EncodeSQL = KmaEncodeSQLDate(ExpressionVariant)
        '                Case FieldTypeLongText, FieldTypeHTML
        '                    EncodeSQL = KmaEncodeSQLLongText(ExpressionVariant)
        '                Case FieldTypeFile, FieldTypeImage, FieldTypeLink, FieldTypeResourceLink, FieldTypeRedirect, FieldTypeManyToMany, FieldTypeText, FieldTypeTextFile, FieldTypeJavascriptFile, FieldTypeXMLFile, FieldTypeCSSFile, FieldTypeHTMLFile
        '                    EncodeSQL = KmaEncodeSQLText(ExpressionVariant)
        '                Case Else
        '                    EncodeSQL = KmaEncodeSQLText(ExpressionVariant)
        '                    On Error GoTo 0
        '                    Call Err.Raise(KmaErrorBase, App.EXEName, "Unknown Field Type [" & fieldType & "] used FieldTypeText.")
        '            End Select
        '            '
        '        End Function
        '        ''
        '        ''
        '        ''
        '        'Public Sub AppendLogFile(Text)
        '        '    On Error GoTo 0
        '        '    Dim kmafs As New kmaFileSystem3.FileSystemClass
        '        '    Dim Filename As String
        '        '    Filename = GetProgramPath() & "\logs\Trace" & kmaEncodeText(CLng(Int(Now()))) & ".log"
        '        '    Call kmafs.AppendLogFile2(Filename, """" & FormatDateTime(Now(), vbGeneralDate) & """,""" & Text & """" & vbCrLf)
        '        '    End Sub
        '        ''
        '        ''========================================================================
        '        ''   HandleError
        '        ''       Logs the error and either resumes next, or raises it to the next level
        '        ''========================================================================
        '        ''
        '        'Public Function HandleError(ClassName As String, MethodName As String, ErrNumber as integer, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNext As Boolean, Optional URL As String) As String
        '        '    ' ##### removed to catch err<>0 problem on error resume next
        '        '    '
        '        '    Dim ErrorMessage As String
        '        '    '
        '        '    If ErrorTrap Then
        '        '        ErrorMessage = ErrorMessage & " Unexpected ErrorTrap"
        '        '    Else
        '        '        ErrorMessage = ErrorMessage & " Error"
        '        '        End If
        '        '    '
        '        '    If URL <> "" Then
        '        '        ErrorMessage = ErrorMessage & " on page [" & URL & "]"
        '        '        End If
        '        '    '
        '        '    If ErrorTrap Then
        '        '        If ResumeNext Then
        '        '            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        '        '        Else
        '        '            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        '        '            On Error GoTo 0
        '        '            Call Err.Raise(ErrNumber, ErrSource, ErrDescription)
        '        '            End If
        '        '    Else
        '        '        If ResumeNext Then
        '        '            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging  [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        '        '        Else
        '        '            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
        '        '            On Error GoTo 0
        '        '            Call Err.Raise(ErrNumber, ErrSource, ErrDescription, , -1)
        '        '            End If
        '        '        End If
        '        '    '
        '        '    End Function
        '        '
        '        '
        '        '
        '        Public Sub cpTick(Text As String)
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            Dim iText As String
        '            Dim Duration as integer
        '            If CPTickCountBase <> 0 Then
        '                Duration = (GetTickCount - CPTickCountBase)
        '            End If
        '            iText = "cpTick " & Format(Duration / 1000, "0000.000") & " " & App.EXEName & " " & Text
        '            Call AppendLogFile(App.EXEName & ".?.cpTick, " & iText)
        '            CPTickCountBase = GetTickCount
        '            '
        '        End Sub
        '        '
        '        '=====================================================================================================
        '        '   Set a value in a name/value pair
        '        '=====================================================================================================
        '        '
        '        Public Sub SetNameValueArrays(InputName As String, InputValue As String, SQLName() As String, SQLValue() As String, Index as integer)
        '            ' ##### removed to catch err<>0 problem on error resume next
        '            '
        '            SQLName(Index) = InputName
        '            SQLValue(Index) = InputValue
        '            Index = Index + 1
        '            '
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Function GetApplicationStatusMessage(ApplicationStatus as integer) As String
        '            Select Case ApplicationStatus
        '                Case ApplicationStatusNoHostService
        '                    GetApplicationStatusMessage = "Contensive server not running"
        '                Case ApplicationStatusNotFound
        '                    GetApplicationStatusMessage = "Contensive application not found"
        '                Case ApplicationStatusLoadedNotRunning
        '                    GetApplicationStatusMessage = "Contensive application not running"
        '                Case ApplicationStatusRunning
        '                    GetApplicationStatusMessage = "Contensive application running"
        '                Case ApplicationStatusStarting
        '                    GetApplicationStatusMessage = "Contensive application starting"
        '                Case ApplicationStatusUpgrading
        '                    GetApplicationStatusMessage = "Contensive database upgrading"
        '                Case ApplicationStatusDbBad
        '                    GetApplicationStatusMessage = "Error verifying core database records"
        '                Case ApplicationStatusDbFailure
        '                    GetApplicationStatusMessage = "Error opening application database"
        '                Case ApplicationStatusKernelFailure
        '                    GetApplicationStatusMessage = "Error contacting Contensive kernel services"
        '                Case ApplicationStatusLicenseFailure
        '                    GetApplicationStatusMessage = "Error verifying Contensive site license, see Http://www.Contensive.com/License"
        '                Case ApplicationStatusConnectionObjectFailure
        '                    GetApplicationStatusMessage = "Error creating ODBC Connection object"
        '                Case ApplicationStatusConnectionStringFailure
        '                    GetApplicationStatusMessage = "ODBC Data Source connection failed"
        '                Case ApplicationStatusDataSourceFailure
        '                    GetApplicationStatusMessage = "Error opening default data source"
        '                Case ApplicationStatusDuplicateDomains
        '                    GetApplicationStatusMessage = "Can not determine application because there are multiple applications with domain names that match this site's domain (See Application Manager)"
        '                Case ApplicationStatusUnknownFailure
        '                    GetApplicationStatusMessage = "Unknown error, see trace log for details (/Contensive/Logs/trace____.log)"
        '                Case ApplicationStatusPaused
        '                    GetApplicationStatusMessage = "Contensive application paused"
        '                Case Else
        '                    GetApplicationStatusMessage = "Unknown status code [" & ApplicationStatus & "], see trace log for details"
        '            End Select
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetFormInputSelectNameValue(SelectName As String, NameValueArray() As NameValuePairType) As String
        '            Dim Pointer as integer
        '            Dim Source() As NameValuePairType
        '            '
        '            Source = NameValueArray
        '            GetFormInputSelectNameValue = "<SELECT name=""" & SelectName & """ Size=""1"">"
        '            For Pointer = 0 To UBound(NameValueArray)
        '                GetFormInputSelectNameValue = GetFormInputSelectNameValue & "<OPTION value=""" & Source(Pointer).Value & """>" & Source(Pointer).Name & "</OPTION>"
        '            Next
        '            GetFormInputSelectNameValue = GetFormInputSelectNameValue & "</SELECT>"
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaGetSpacer(Width as integer, Height as integer) As String
        '            kmaGetSpacer = "<img alt=""space"" src=""/ccLib/images/spacer.gif"" width=""" & Width & """ height=""" & Height & """ border=""0"">"
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function kmaProcessReplacement(NameValueLines As Object, Source As Object) As String
        '            '
        '            Dim iNameValueLines As String
        '            Dim Lines() As String
        '            Dim LineCnt as integer
        '            Dim LinePtr as integer
        '            '
        '            Dim Names() As String
        '            Dim Values() As String
        '            Dim PairPtr as integer
        '            Dim PairCnt as integer
        '            Dim Splits() As String
        '            '
        '            ' ----- read pairs in from NameValueLines
        '            '
        '            iNameValueLines = kmaEncodeText(NameValueLines)
        '            If InStr(1, iNameValueLines, "=") <> 0 Then
        '                PairCnt = 0
        '                Lines = SplitCRLF(iNameValueLines)
        '                For LinePtr = 0 To UBound(Lines)
        '                    If InStr(1, Lines(LinePtr), "=") <> 0 Then
        '                        Splits = Split(Lines(LinePtr), "=")
        '                        ReDim Preserve Names(PairCnt)
        '                        ReDim Preserve Names(PairCnt)
        '                        ReDim Preserve Values(PairCnt)
        '                        Names(PairCnt) = Trim(Splits(0))
        '                        Names(PairCnt) = Replace(Names(PairCnt), vbTab, "")
        '                        Splits(0) = ""
        '                        Values(PairCnt) = Trim(Splits(1))
        '                        PairCnt = PairCnt + 1
        '                    End If
        '                Next
        '            End If
        '            '
        '            ' ----- Process replacements on Source
        '            '
        '            kmaProcessReplacement = kmaEncodeText(Source)
        '            If PairCnt > 0 Then
        '                For PairPtr = 0 To PairCnt - 1
        '                    kmaProcessReplacement = Replace(kmaProcessReplacement, Names(PairPtr), Values(PairPtr), 1, 999, 1)
        '                Next
        '            End If
        '            '
        '        End Function
        '        '
        '        '==========================================================================================================================
        '        '   To convert from site license to server licenses, we still need the URLEncoder in the site license
        '        '   This routine generates a site license that is just the URL encoder.
        '        '==========================================================================================================================
        '        '
        '        Public Function GetURLEncoder() As String
        '            Randomize()
        '            GetURLEncoder = CStr(Int(1 + (Rnd() * 8))) & CStr(Int(1 + (Rnd() * 8))) & CStr(Int(1000000000 + (Rnd() * 899999999)))
        '        End Function
        '        '
        '        '==========================================================================================================================
        '        '   To convert from site license to server licenses, we still need the URLEncoder in the site license
        '        '   This routine generates a site license that is just the URL encoder.
        '        '==========================================================================================================================
        '        '
        '        Public Function GetSiteLicenseKey() As String
        '            GetSiteLicenseKey = "00000-00000-00000-" & GetURLEncoder()
        '        End Function
        '        '
        '        '
        '        '
        'Public Sub ccAddTabEntry(Caption As String, Link As String, IsHit As Boolean, Optional StylePrefix As String, Optional LiveBody As String)
        '            On Error GoTo ErrorTrap
        '            '
        '            If TabsCnt <= TabsSize Then
        '                TabsSize = TabsSize + 10
        '                ReDim Preserve Tabs(TabsSize)
        '            End If
        '            With Tabs(TabsCnt)
        '                .Caption = Caption
        '                .Link = Link
        '                .IsHit = IsHit
        '                .StylePrefix = KmaEncodeMissingText(StylePrefix, "cc")
        '                .LiveBody = KmaEncodeMissingText(LiveBody, "")
        '            End With
        '            TabsCnt = TabsCnt + 1
        '            '
        '            Exit Sub
        '            '
        'ErrorTrap:
        '            Call Err.Raise(Err.Number, Err.Source, "Error in ccAddTabEntry-" & Err.Description)
        '        End Sub
        '        '
        '        '
        '        '
        '        Public Function OldccGetTabs() As String
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim TabPtr as integer
        '            Dim HitPtr as integer
        '            Dim IsLiveTab As Boolean
        '            Dim TabBody As String
        '            Dim TabLink As String
        '            Dim TabID As String
        '            Dim FirstLiveBodyShown As Boolean
        '            '
        '            If TabsCnt > 0 Then
        '                HitPtr = 0
        '                '
        '                ' Create TabBar
        '                '
        '                OldccGetTabs = OldccGetTabs & "<table border=0 cellspacing=0 cellpadding=0 align=center ><tr>"
        '                For TabPtr = 0 To TabsCnt - 1
        '                    TabID = CStr(GetRandomInteger())
        '                    If Tabs(TabPtr).LiveBody = "" Then
        '                        '
        '                        ' This tab is linked to a page
        '                        '
        '                        TabLink = cp.Utils.EncodeHTML(Tabs(TabPtr).Link)
        '                    Else
        '                        '
        '                        ' This tab has a live body
        '                        '
        '                        TabLink = cp.Utils.EncodeHTML(Tabs(TabPtr).Link)
        '                        If Not FirstLiveBodyShown Then
        '                            FirstLiveBodyShown = True
        '                            TabBody = TabBody & "<div style=""visibility: visible; position: absolute; left: 0px;"" class=""" & Tabs(TabPtr).StylePrefix & "Body"" id=""" & TabID & """></div>"
        '                        Else
        '                            TabBody = TabBody & "<div style=""visibility: hidden; position: absolute; left: 0px;"" class=""" & Tabs(TabPtr).StylePrefix & "Body"" id=""" & TabID & """></div>"
        '                        End If
        '                    End If
        '                    OldccGetTabs = OldccGetTabs & "<td valign=bottom>"
        '                    If Tabs(TabPtr).IsHit And (HitPtr = 0) Then
        '                        HitPtr = TabPtr
        '                        '
        '                        ' This tab is hit
        '                        '
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<table cellspacing=0 cellPadding=0 border=0>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=2 height=1 width=2></td>" _
        '                            & "<td colspan=1 height=1 bgcolor=black></td>" _
        '                            & "<td colspan=3 height=1 width=3></td>" _
        '                            & "</tr>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=1 height=1 width=1></td>" _
        '                            & "<td colspan=1 height=1 width=1 bgcolor=black></td>" _
        '                            & "<td colspan=1 height=1></td>" _
        '                            & "<td colspan=1 height=1 width=1 bgcolor=black></td>" _
        '                            & "<td colspan=2 height=1 width=2></td>" _
        '                            & "</tr>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=1 height=2 bgcolor=black></td>" _
        '                            & "<td colspan=1 height=2></td>" _
        '                            & "<td colspan=1 height=2></td>" _
        '                            & "<td colspan=1 height=2></td>" _
        '                            & "<td colspan=1 height=2 width=1 bgcolor=black></td>" _
        '                            & "<td colspan=1 height=2 width=1></td>" _
        '                            & "</tr>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "<td></td>" _
        '                            & "<td>" _
        '                            & "<table cellspacing=0 cellPadding=2 border=0><tr>" _
        '                            & "<td Class=""ccTabHit"">&nbsp;<a href=""" & TabLink & """ Class=""ccTabHit"">" & Tabs(TabPtr).Caption & "</a>&nbsp;</td>" _
        '                            & "</tr></table >" _
        '                            & "</td>" _
        '                            & "<td></td>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "<td></td>" _
        '                            & "</tr>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "<td></td>" _
        '                            & "<td></td>" _
        '                            & "<td></td>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "</tr>" _
        '                            & "</table >"
        '                    Else
        '                        '
        '                        ' This tab is not hit
        '                        '
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<table cellspacing=0 cellPadding=0 border=0>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=6 height=1></td>" _
        '                            & "</tr>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=2 height=1></td>" _
        '                            & "<td colspan=1 height=1 bgcolor=black></td>" _
        '                            & "<td colspan=3 height=1></td>" _
        '                            & "</tr>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td width=1></td>" _
        '                            & "<td width=1 bgcolor=black></td>" _
        '                            & "<td></td>" _
        '                            & "<td width=1 bgcolor=black></td>" _
        '                            & "<td width=2 colspan=2></td>" _
        '                            & "</tr>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td width=1 bgcolor=black></td>" _
        '                            & "<td width=1></td>" _
        '                            & "<td nowrap>" _
        '                            & "<table cellspacing=0 cellPadding=2 border=0><tr>" _
        '                            & "<td Class=""ccTab"">&nbsp;<a href=""" & TabLink & """ Class=""ccTab"">" & Tabs(TabPtr).Caption & "</a>&nbsp;</td>" _
        '                            & "</tr></table >" _
        '                            & "</td>" _
        '                            & "<td width=1></td>" _
        '                            & "<td width=1 bgcolor=black></td>" _
        '                            & "<td width=1></td>" _
        '                            & "</tr>"
        '                        OldccGetTabs = OldccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=6 height=1 bgcolor=black></td>" _
        '                            & "</tr>" _
        '                            & "</table >"
        '                    End If
        '                    OldccGetTabs = OldccGetTabs & "</td>"
        '                Next
        '                OldccGetTabs = OldccGetTabs & "<td class=""ccTabEnd"">&nbsp;</td></tr>"
        '                If TabBody <> "" Then
        '                    OldccGetTabs = OldccGetTabs & "<tr><td colspan=6>" & TabBody & "</td></tr>"
        '                End If
        '                OldccGetTabs = OldccGetTabs & "</tr></table >"
        '                TabsCnt = 0
        '            End If
        '            '
        '            Exit Function
        '            '
        'ErrorTrap:
        '            Call Err.Raise(Err.Number, Err.Source, "Error in OldccGetTabs-" & Err.Description)
        '        End Function


        '        '
        '        '
        '        '
        '        Public Function ccGetTabs() As String
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim TabPtr as integer
        '            Dim HitPtr as integer
        '            Dim IsLiveTab As Boolean
        '            Dim TabBody As String
        '            Dim TabLink As String
        '            Dim TabID As String
        '            Dim FirstLiveBodyShown As Boolean
        '            '
        '            If TabsCnt > 0 Then
        '                HitPtr = 0
        '                '
        '                ' Create TabBar
        '                '
        '                ccGetTabs = ccGetTabs & "<table border=0 cellspacing=0 cellpadding=0 align=center ><tr>"
        '                For TabPtr = 0 To TabsCnt - 1
        '                    TabID = CStr(GetRandomInteger())
        '                    If Tabs(TabPtr).LiveBody = "" Then
        '                        '
        '                        ' This tab is linked to a page
        '                        '
        '                        TabLink = cp.Utils.EncodeHTML(Tabs(TabPtr).Link)
        '                    Else
        '                        '
        '                        ' This tab has a live body
        '                        '
        '                        TabLink = cp.Utils.EncodeHTML(Tabs(TabPtr).Link)
        '                        If Not FirstLiveBodyShown Then
        '                            FirstLiveBodyShown = True
        '                            TabBody = TabBody & "<div style=""visibility: visible; position: absolute; left: 0px;"" class=""" & Tabs(TabPtr).StylePrefix & "Body"" id=""" & TabID & """>" & Tabs(TabPtr).LiveBody & "</div>"
        '                        Else
        '                            TabBody = TabBody & "<div style=""visibility: hidden; position: absolute; left: 0px;"" class=""" & Tabs(TabPtr).StylePrefix & "Body"" id=""" & TabID & """>" & Tabs(TabPtr).LiveBody & "</div>"
        '                        End If
        '                    End If
        '                    ccGetTabs = ccGetTabs & "<td valign=bottom>"
        '                    If Tabs(TabPtr).IsHit And (HitPtr = 0) Then
        '                        HitPtr = TabPtr
        '                        '
        '                        ' This tab is hit
        '                        '
        '                        ccGetTabs = ccGetTabs _
        '                            & "<table cellspacing=0 cellPadding=0 border=0>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=2 height=1 width=2></td>" _
        '                            & "<td colspan=1 height=1 bgcolor=black></td>" _
        '                            & "<td colspan=3 height=1 width=3></td>" _
        '                            & "</tr>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=1 height=1 width=1></td>" _
        '                            & "<td colspan=1 height=1 width=1 bgcolor=black></td>" _
        '                            & "<td Class=""ccTabHit"" colspan=1 height=1></td>" _
        '                            & "<td colspan=1 height=1 width=1 bgcolor=black></td>" _
        '                            & "<td colspan=2 height=1 width=2></td>" _
        '                            & "</tr>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=1 height=2 bgcolor=black></td>" _
        '                            & "<td Class=""ccTabHit"" colspan=1 height=2></td>" _
        '                            & "<td Class=""ccTabHit"" colspan=1 height=2></td>" _
        '                            & "<td Class=""ccTabHit"" colspan=1 height=2></td>" _
        '                            & "<td colspan=1 height=2 bgcolor=black></td>" _
        '                            & "<td colspan=1 height=2></td>" _
        '                            & "</tr>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "<td Class=""ccTabHit""></td>" _
        '                            & "<td Class=""ccTabHit"">" _
        '                            & "<table cellspacing=0 cellPadding=2 border=0><tr>" _
        '                            & "<td Class=""ccTabHit"">&nbsp;<a href=""" & TabLink & """ Class=""ccTabHit"">" & Tabs(TabPtr).Caption & "</a>&nbsp;</td>" _
        '                            & "</tr></table >" _
        '                            & "</td>" _
        '                            & "<td Class=""ccTabHit""></td>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "<td></td>" _
        '                            & "</tr>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "<td Class=""ccTabHit""></td>" _
        '                            & "<td Class=""ccTabHit""></td>" _
        '                            & "<td Class=""ccTabHit""></td>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "<td bgcolor=black></td>" _
        '                            & "</tr>" _
        '                            & "</table >"
        '                    Else
        '                        '
        '                        ' This tab is not hit
        '                        '
        '                        ccGetTabs = ccGetTabs _
        '                            & "<table cellspacing=0 cellPadding=0 border=0>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=6 height=1></td>" _
        '                            & "</tr>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=2 height=1></td>" _
        '                            & "<td colspan=1 height=1 bgcolor=black></td>" _
        '                            & "<td colspan=3 height=1></td>" _
        '                            & "</tr>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td width=1></td>" _
        '                            & "<td width=1 bgcolor=black></td>" _
        '                            & "<td Class=""ccTab""></td>" _
        '                            & "<td width=1 bgcolor=black></td>" _
        '                            & "<td width=2 colspan=2></td>" _
        '                            & "</tr>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td width=1 bgcolor=black></td>" _
        '                            & "<td width=1 Class=""ccTab""></td>" _
        '                            & "<td nowrap Class=""ccTab"">" _
        '                            & "<table cellspacing=0 cellPadding=2 border=0><tr>" _
        '                            & "<td Class=""ccTab"">&nbsp;<a href=""" & TabLink & """ Class=""ccTab"">" & Tabs(TabPtr).Caption & "</a>&nbsp;</td>" _
        '                            & "</tr></table >" _
        '                            & "</td>" _
        '                            & "<td width=1 Class=""ccTab""></td>" _
        '                            & "<td width=1 bgcolor=black></td>" _
        '                            & "<td width=1></td>" _
        '                            & "</tr>"
        '                        ccGetTabs = ccGetTabs _
        '                            & "<tr>" _
        '                            & "<td colspan=6 height=1 bgcolor=black></td>" _
        '                            & "</tr>" _
        '                            & "</table >"
        '                    End If
        '                    ccGetTabs = ccGetTabs & "</td>"
        '                Next
        '                ccGetTabs = ccGetTabs & "<td class=""ccTabEnd"">&nbsp;</td></tr>"
        '                If TabBody <> "" Then
        '                    ccGetTabs = ccGetTabs & "<tr><td colspan=6>" & TabBody & "</td></tr>"
        '                End If
        '                ccGetTabs = ccGetTabs & "</tr></table >"
        '                TabsCnt = 0
        '            End If
        '            '
        '            Exit Function
        '            '
        'ErrorTrap:
        '            Call Err.Raise(Err.Number, Err.Source, "Error in ccGetTabs-" & Err.Description)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ConvertLinksToAbsolute(Source As String, RootLink As String) As String
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim s As String
        '            '
        '            s = Source
        '            '
        '            s = Replace(s, " href=""", " href=""/", , , vbTextCompare)
        '            s = Replace(s, " href=""/http", " href=""http", , , vbTextCompare)
        '            s = Replace(s, " href=""/mailto", " href=""mailto", , , vbTextCompare)
        '            s = Replace(s, " href=""//", " href=""" & RootLink, , , vbTextCompare)
        '            s = Replace(s, " href=""/?", " href=""" & RootLink & "?", , , vbTextCompare)
        '            s = Replace(s, " href=""/", " href=""" & RootLink, , , vbTextCompare)
        '            '
        '            s = Replace(s, " href=", " href=/", , , vbTextCompare)
        '            s = Replace(s, " href=/""", " href=""", , , vbTextCompare)
        '            s = Replace(s, " href=/http", " href=http", , , vbTextCompare)
        '            s = Replace(s, " href=//", " href=" & RootLink, , , vbTextCompare)
        '            s = Replace(s, " href=/?", " href=" & RootLink & "?", , , vbTextCompare)
        '            s = Replace(s, " href=/", " href=" & RootLink, , , vbTextCompare)
        '            '
        '            s = Replace(s, " src=""", " src=""/", , , vbTextCompare)
        '            s = Replace(s, " src=""/http", " src=""http", , , vbTextCompare)
        '            s = Replace(s, " src=""//", " src=""" & RootLink, , , vbTextCompare)
        '            s = Replace(s, " src=""/?", " src=""" & RootLink & "?", , , vbTextCompare)
        '            s = Replace(s, " src=""/", " src=""" & RootLink, , , vbTextCompare)
        '            '
        '            s = Replace(s, " src=", " src=/", , , vbTextCompare)
        '            s = Replace(s, " src=/""", " src=""", , , vbTextCompare)
        '            s = Replace(s, " src=/http", " src=http", , , vbTextCompare)
        '            s = Replace(s, " src=//", " src=" & RootLink, , , vbTextCompare)
        '            s = Replace(s, " src=/?", " src=" & RootLink & "?", , , vbTextCompare)
        '            s = Replace(s, " src=/", " src=" & RootLink, , , vbTextCompare)
        '            '
        '            ConvertLinksToAbsolute = s
        '            '
        '            Exit Function
        '            '
        'ErrorTrap:
        '            Call Err.Raise(Err.Number, Err.Source, "Error in ConvertLinksToAbsolute-" & Err.Description)
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetProgramPath() As String
        '            GetProgramPath = App.Path
        '            If InStr(1, GetProgramPath, "c:\h\contensive\", vbTextCompare) <> 0 Then
        '                GetProgramPath = "c:\h\Contensive"
        '            ElseIf InStr(1, GetProgramPath, "c:\release\", vbTextCompare) <> 0 Then
        '                GetProgramPath = "c:\h\Contensive"
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetAddonRootPath() As String
        '            GetAddonRootPath = GetProgramPath()
        '            If InStr(1, GetAddonRootPath, "c:\h\contensive", vbTextCompare) <> 0 Then
        '                '
        '                ' debugging - change program path to dummy path so addon builds all copy to
        '                '
        '                GetAddonRootPath = "c:\program files\kma\contensive"
        '            End If
        '            GetAddonRootPath = GetAddonRootPath & "\addons"
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetHTMLComment(Comment) As String
        '            GetHTMLComment = "<!-- " & Comment & " -->"
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function SplitCRLF(Expression As String) As String()
        '            Dim Args() As String
        '            Dim Ptr as integer
        '            '
        '            If InStr(1, Expression, vbCrLf) <> 0 Then
        '                SplitCRLF = Split(Expression, vbCrLf, , vbTextCompare)
        '            ElseIf InStr(1, Expression, vbCr) <> 0 Then
        '                SplitCRLF = Split(Expression, vbCr, , vbTextCompare)
        '            ElseIf InStr(1, Expression, vbLf) <> 0 Then
        '                SplitCRLF = Split(Expression, vbLf, , vbTextCompare)
        '            Else
        '                ReDim SplitCRLF(0)
        '                SplitCRLF = Split(Expression, vbCrLf)
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        'Public Sub kmaShell(Cmd As String, Optional ByVal eWindowStyle As VBA.VbAppWinStyle = vbHide, Optional WaitForReturn As Boolean)
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim ShellObj As Object
        '            '
        '            ShellObj = CreateObject("WScript.Shell")
        '            If Not (ShellObj Is Nothing) Then
        '                Call ShellObj.Run(Cmd, 0, WaitForReturn)
        '            End If
        '            ShellObj = Nothing
        '            Exit Sub
        '            '
        'ErrorTrap:
        '            Call AppendLogFile("ErrorTrap, kmaShell running command [" & Cmd & "], WaitForReturn=" & WaitForReturn & ", err=" & GetErrString(Err))
        '        End Sub
        '        '
        '        '------------------------------------------------------------------------------------------------------------
        '        '   Encodes an argument in an Addon OptionString (QueryString) for all non-allowed characters
        '        '       call this before parsing them together
        '        '       call decodeAddonConstructorArgument after parsing them apart
        '        '
        '        '       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=VAlue[Option1|Option2]
        '        '
        '        '       This routine is needed for all Arg, Name, Value, Option values
        '        '
        '        '------------------------------------------------------------------------------------------------------------
        '        '
        '        Public Function EncodeAddonConstructorArgument(Arg As String) As String
        '            Dim a As String
        '            If Arg <> "" Then
        '                a = Arg
        '                If True Then
        '                    'If AddonNewParse Then
        '                    a = Replace(a, "\", "\\")
        '                    a = Replace(a, vbCrLf, "\n")
        '                    a = Replace(a, vbTab, "\t")
        '                    a = Replace(a, "&", "\&")
        '                    a = Replace(a, "=", "\=")
        '                    a = Replace(a, ",", "\,")
        '                    a = Replace(a, """", "\""")
        '                    a = Replace(a, "'", "\'")
        '                    a = Replace(a, "|", "\|")
        '                    a = Replace(a, "[", "\[")
        '                    a = Replace(a, "]", "\]")
        '                    a = Replace(a, ":", "\:")
        '                End If
        '                EncodeAddonConstructorArgument = a
        '            End If
        '        End Function
        '        '
        '        '------------------------------------------------------------------------------------------------------------
        '        '   Decodes an argument parsed from an AddonConstructorString for all non-allowed characters
        '        '       AddonConstructorString is a & delimited string of name=value[selector]descriptor
        '        '
        '        '       to get a value from an AddonConstructorString, first use getargument() to get the correct value[selector]descriptor
        '        '       then remove everything to the right of any '['
        '        '
        '        '       call encodeAddonConstructorargument before parsing them together
        '        '       call decodeAddonConstructorArgument after parsing them apart
        '        '
        '        '       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=VAlue[Option1|Option2]
        '        '
        '        '       This routine is needed for all Arg, Name, Value, Option values
        '        '
        '        '------------------------------------------------------------------------------------------------------------
        '        '
        '        Public Function DecodeAddonConstructorArgument(EncodedArg As String) As String
        '            Dim a As String
        '            Dim Pos as integer
        '            '
        '            a = EncodedArg
        '            If True Then
        '                'If AddonNewParse Then
        '                a = Replace(a, "\:", ":")
        '                a = Replace(a, "\]", "]")
        '                a = Replace(a, "\[", "[")
        '                a = Replace(a, "\|", "|")
        '                a = Replace(a, "\'", "'")
        '                a = Replace(a, "\""", """")
        '                a = Replace(a, "\,", ",")
        '                a = Replace(a, "\=", "=")
        '                a = Replace(a, "\&", "&")
        '                a = Replace(a, "\t", vbTab)
        '                a = Replace(a, "\n", vbCrLf)
        '                a = Replace(a, "\\", "\")
        '            End If
        '            DecodeAddonConstructorArgument = a
        '        End Function
        '        '
        '        '------------------------------------------------------------------------------------------------------------
        '        '   use only internally
        '        '
        '        '   encode an argument to be used in a name=value& (N-V-A) string
        '        '
        '        '   an argument can be any one of these is this format:
        '        '       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=Value[Option1|Option2]descriptor
        '        '
        '        '   to create an nva string
        '        '       string = encodeNvaArgument( name ) & "=" & encodeNvaArgument( value ) & "&"
        '        '
        '        '   to decode an nva string
        '        '       split on ampersand then on equal, and decodeNvaArgument() each part
        '        '
        '        '------------------------------------------------------------------------------------------------------------
        '        '
        '        Public Function encodeNvaArgument(Arg As String) As String
        '            Dim a As String
        '            a = Arg
        '            If a <> "" Then
        '                a = Replace(a, vbCrLf, "#0013#")
        '                a = Replace(a, vbLf, "#0013#")
        '                a = Replace(a, vbCr, "#0013#")
        '                a = Replace(a, "&", "#0038#")
        '                a = Replace(a, "=", "#0061#")
        '                a = Replace(a, ",", "#0044#")
        '                a = Replace(a, """", "#0034#")
        '                a = Replace(a, "'", "#0039#")
        '                a = Replace(a, "|", "#0124#")
        '                a = Replace(a, "[", "#0091#")
        '                a = Replace(a, "]", "#0093#")
        '                a = Replace(a, ":", "#0058#")
        '            End If
        '            encodeNvaArgument = a
        '        End Function
        '        '
        '        '------------------------------------------------------------------------------------------------------------
        '        '   use only internally
        '        '       decode an argument removed from a name=value& string
        '        '       see encodeNvaArgument for details on how to use this
        '        '------------------------------------------------------------------------------------------------------------
        '        '
        '        Public Function decodeNvaArgument(EncodedArg As String) As String
        '            Dim a As String
        '            '
        '            a = EncodedArg
        '            a = Replace(a, "#0058#", ":")
        '            a = Replace(a, "#0093#", "]")
        '            a = Replace(a, "#0091#", "[")
        '            a = Replace(a, "#0124#", "|")
        '            a = Replace(a, "#0039#", "'")
        '            a = Replace(a, "#0034#", """")
        '            a = Replace(a, "#0044#", ",")
        '            a = Replace(a, "#0061#", "=")
        '            a = Replace(a, "#0038#", "&")
        '            a = Replace(a, "#0013#", vbCrLf)
        '            decodeNvaArgument = a
        '        End Function
        '        '
        '        ' returns true of the link is a valid link on the source host
        '        '
        '        Public Function IsLinkToThisHost(Host As String, Link As String) As Boolean
        '            '
        '            Dim LinkHost As String
        '            Dim Pos as integer
        '            '
        '            If Trim(Link) = "" Then
        '                '
        '                ' Blank is not a link
        '                '
        '                IsLinkToThisHost = False
        '            ElseIf InStr(1, Link, "://") <> 0 Then
        '                '
        '                ' includes protocol, may be link to another site
        '                '
        '                LinkHost = LCase(Link)
        '                Pos = 1
        '                Pos = InStr(Pos, LinkHost, "://")
        '                If Pos > 0 Then
        '                    Pos = InStr(Pos + 3, LinkHost, "/")
        '                    If Pos > 0 Then
        '                        LinkHost = Mid(LinkHost, 1, Pos - 1)
        '                    End If
        '                    IsLinkToThisHost = (LCase(Host) = LinkHost)
        '                    If Not IsLinkToThisHost Then
        '                        '
        '                        ' try combinations including/excluding www.
        '                        '
        '                        If InStr(1, LinkHost, "www.", vbTextCompare) <> 0 Then
        '                            '
        '                            ' remove it
        '                            '
        '                            LinkHost = Replace(LinkHost, "www.", "", 1, -1, vbTextCompare)
        '                            IsLinkToThisHost = (LCase(Host) = LinkHost)
        '                        Else
        '                            '
        '                            ' add it
        '                            '
        '                            LinkHost = Replace(LinkHost, "://", "://www.", 1, -1, vbTextCompare)
        '                            IsLinkToThisHost = (LCase(Host) = LinkHost)
        '                        End If
        '                    End If
        '                End If
        '            ElseIf InStr(1, Link, "#") = 1 Then
        '                '
        '                ' Is a bookmark, not a link
        '                '
        '                IsLinkToThisHost = False
        '            Else
        '                '
        '                ' all others are links on the source
        '                '
        '                IsLinkToThisHost = True
        '            End If
        '            If Not IsLinkToThisHost Then
        '                Link = Link
        '            End If
        '        End Function
        '        '
        '        '========================================================================================================
        '        '   ConvertLinkToRootRelative
        '        '
        '        '   /images/logo-main.jpg with any Basepath to /images/logo-main.jpg
        '        '   http://gcm.brandeveolve.com/images/logo-main.jpg with any BasePath  to /images/logo-main.jpg
        '        '   images/logo-main.jpg with Basepath '/' to /images/logo-main.jpg
        '        '   logo-main.jpg with Basepath '/images/' to /images/logo-main.jpg
        '        '
        '        '========================================================================================================
        '        '
        '        Public Function ConvertLinkToRootRelative(Link As String, BasePath As String) As String
        '            '
        '            Dim Pos as integer
        '            '
        '            ConvertLinkToRootRelative = Link
        '            If InStr(1, Link, "/") = 1 Then
        '                '
        '                '   case /images/logo-main.jpg with any Basepath to /images/logo-main.jpg
        '                '
        '            ElseIf InStr(1, Link, "://") <> 0 Then
        '                '
        '                '   case http://gcm.brandeveolve.com/images/logo-main.jpg with any BasePath  to /images/logo-main.jpg
        '                '
        '                Pos = InStr(1, Link, "://")
        '                If Pos > 0 Then
        '                    Pos = InStr(Pos + 3, Link, "/")
        '                    If Pos > 0 Then
        '                        ConvertLinkToRootRelative = Mid(Link, Pos)
        '                    Else
        '                        '
        '                        ' This is just the domain name, RootRelative is the root
        '                        '
        '                        ConvertLinkToRootRelative = "/"
        '                    End If
        '                End If
        '            Else
        '                '
        '                '   case images/logo-main.jpg with Basepath '/' to /images/logo-main.jpg
        '                '   case logo-main.jpg with Basepath '/images/' to /images/logo-main.jpg
        '                '
        '                ConvertLinkToRootRelative = BasePath & Link
        '            End If
        '            '
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function GetAddonIconImg(AdminURL As String, IconWidth as integer, IconHeight as integer, IconSprites as integer, IconIsInline As Boolean, IconImgID As String, IconFilename As String, serverFilePath As String, IconAlt As String, IconTitle As String, ACInstanceID As String, IconSpriteColumn as integer) As String
        '            '
        '            Dim ImgStyle As String
        '            Dim IconHeightNumeric as integer
        '            '
        '            If IconAlt = "" Then
        '                IconAlt = "Add-on"
        '            End If
        '            If IconTitle = "" Then
        '                IconTitle = "Rendered as Add-on"
        '            End If
        '            If IconFilename = "" Then
        '                '
        '                ' No icon given, use the default
        '                '
        '                If IconIsInline Then
        '                    IconFilename = "/ccLib/images/IconAddonInlineDefault.png"
        '                    IconWidth = 62
        '                    IconHeight = 17
        '                    IconSprites = 0
        '                Else
        '                    IconFilename = "/ccLib/images/IconAddonBlockDefault.png"
        '                    IconWidth = 57
        '                    IconHeight = 59
        '                    IconSprites = 4
        '                End If
        '            ElseIf InStr(1, IconFilename, "://") <> 0 Then
        '                '
        '                ' icon is an Absolute URL - leave it
        '                '
        '            ElseIf Left(IconFilename, 1) = "/" Then
        '                '
        '                ' icon is Root Relative, leave it
        '                '
        '            Else
        '                '
        '                ' icon is a virtual file, add the serverfilepath
        '                '
        '                IconFilename = serverFilePath & IconFilename
        '            End If
        '            'IconFilename = kmaEncodeJavascript(IconFilename)
        '            If (IconWidth = 0) Or (IconHeight = 0) Then
        '                IconSprites = 0
        '            End If

        '            If IconSprites = 0 Then
        '                '
        '                ' just the icon
        '                '
        '                GetAddonIconImg = "<img" _
        '                    & " border=0" _
        '                    & " id=""" & IconImgID & """" _
        '                    & " onDblClick=""window.parent.OpenAddonPropertyWindow(this,'" & AdminURL & "');""" _
        '                    & " alt=""" & IconAlt & """" _
        '                    & " title=""" & IconTitle & """" _
        '                    & " src=""" & IconFilename & """"
        '                'GetAddonIconImg = "<img" _
        '                '    & " id=""AC,AGGREGATEFUNCTION,0," & FieldName & "," & ArgumentList & """" _
        '                '    & " onDblClick=""window.parent.OpenAddonPropertyWindow(this);""" _
        '                '    & " alt=""" & IconAlt & """" _
        '                '    & " title=""" & IconTitle & """" _
        '                '    & " src=""" & IconFilename & """"
        '                If IconWidth <> 0 Then
        '                    GetAddonIconImg = GetAddonIconImg & " width=""" & IconWidth & "px"""
        '                End If
        '                If IconHeight <> 0 Then
        '                    GetAddonIconImg = GetAddonIconImg & " height=""" & IconHeight & "px"""
        '                End If
        '                If IconIsInline Then
        '                    GetAddonIconImg = GetAddonIconImg & " style=""vertical-align:middle;display:inline;"" "
        '                Else
        '                    GetAddonIconImg = GetAddonIconImg & " style=""display:block"" "
        '                End If
        '                If ACInstanceID <> "" Then
        '                    GetAddonIconImg = GetAddonIconImg & " ACInstanceID=""" & ACInstanceID & """"
        '                End If
        '                GetAddonIconImg = GetAddonIconImg & ">"
        '            Else
        '                '
        '                ' Sprite Icon
        '                '
        '                GetAddonIconImg = GetIconSprite(IconImgID, IconSpriteColumn, IconFilename, IconWidth, IconHeight, IconAlt, IconTitle, "window.parent.OpenAddonPropertyWindow(this,'" & AdminURL & "');", IconIsInline, ACInstanceID)
        '                '        GetAddonIconImg = "<img" _
        '                '            & " border=0" _
        '                '            & " id=""" & IconImgID & """" _
        '                '            & " onMouseOver=""this.style.backgroundPosition='" & (-1 * IconSpriteColumn * IconWidth) & "px -" & (2 * IconHeight) & "px'""" _
        '                '            & " onMouseOut=""this.style.backgroundPosition='" & (-1 * IconSpriteColumn * IconWidth) & "px 0px'""" _
        '                '            & " onDblClick=""window.parent.OpenAddonPropertyWindow(this,'" & AdminURL & "');""" _
        '                '            & " alt=""" & IconAlt & """" _
        '                '            & " title=""" & IconTitle & """" _
        '                '            & " src=""/ccLib/images/spacer.gif"""
        '                '        ImgStyle = "background:url(" & IconFilename & ") " & (-1 * IconSpriteColumn * IconWidth) & "px 0px no-repeat;"
        '                '        ImgStyle = ImgStyle & "width:" & IconWidth & "px;"
        '                '        ImgStyle = ImgStyle & "height:" & IconHeight & "px;"
        '                '        If IconIsInline Then
        '                '            'GetAddonIconImg = GetAddonIconImg & " align=""middle"""
        '                '            ImgStyle = ImgStyle & "vertical-align:middle;display:inline;"
        '                '        Else
        '                '            ImgStyle = ImgStyle & "display:block;"
        '                '        End If
        '                '
        '                '
        '                '        'Return_IconStyleMenuEntries = Return_IconStyleMenuEntries & vbCrLf & ",["".icon" & AddonID & """,false,"".icon" & AddonID & """,""background:url(" & IconFilename & ") 0px 0px no-repeat;"
        '                '        'GetAddonIconImg = "<img" _
        '                '        '    & " id=""AC,AGGREGATEFUNCTION,0," & FieldName & "," & ArgumentList & """" _
        '                '        '    & " onMouseOver=""this.style.backgroundPosition=\'0px -" & (2 * IconHeight) & "px\'""" _
        '                '        '    & " onMouseOut=""this.style.backgroundPosition=\'0px 0px\'""" _
        '                '        '    & " onDblClick=""window.parent.OpenAddonPropertyWindow(this);""" _
        '                '        '    & " alt=""" & IconAlt & """" _
        '                '        '    & " title=""" & IconTitle & """" _
        '                '        '    & " src=""/ccLib/images/spacer.gif"""
        '                '        If ACInstanceID <> "" Then
        '                '            GetAddonIconImg = GetAddonIconImg & " ACInstanceID=""" & ACInstanceID & """"
        '                '        End If
        '                '        GetAddonIconImg = GetAddonIconImg & " style=""" & ImgStyle & """>"
        '                '        'Return_IconStyleMenuEntries = Return_IconStyleMenuEntries & """]"
        '            End If
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function ConvertRSTypeToGoogleType(RecordFieldType as integer) As String
        '            Select Case RecordFieldType
        '                Case 2, 3, 4, 5, 6, 14, 16, 17, 18, 19, 20, 21, 131
        '                    ConvertRSTypeToGoogleType = "number"
        '                Case Else
        '                    ConvertRSTypeToGoogleType = "string"
        '            End Select
        '        End Function

        '        '
        '        '========================================================================
        '        '   HandleError
        '        '       Logs the error and either resumes next, or raises it to the next level
        '        '========================================================================
        '        '
        '        Public Sub AppendLogFile2(ContensiveAppName As String, Context As String, ProgramName As String, ClassName As String, MethodName As String, ErrNumber as integer, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNextAfterLogging As Boolean, URL As String, LogFolder As String, LogNamePrefix As String)
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim MonthNumber as integer
        '            Dim DayNumber as integer
        '            Dim FilenameNoExt As String
        '            Dim kmafs As Object
        '            Dim ErrorMessage As String
        '            Dim LogLine As String
        '            Dim ResumeMessage As String
        '            Dim FolderFileList As String
        '            Dim FolderFiles() As String
        '            Dim Ptr as integer
        '            Dim PathFilenameNoExt As String
        '            Dim FileDetails() As String
        '            Dim fileSize as integer
        '            Dim RetryCnt as integer
        '            Dim SaveOK As Boolean
        '            Dim FileSuffix As String
        '            Dim iLogFolder As String
        '            '
        '            iLogFolder = LogFolder
        '            '
        '            If ErrorTrap Then
        '                ErrorMessage = "Error Trap"
        '            Else
        '                ErrorMessage = "Log Entry"
        '            End If
        '            '
        '            If ResumeNextAfterLogging Then
        '                ResumeMessage = "Resume after logging"
        '            Else
        '                ResumeMessage = "Abort after logging"
        '            End If
        '            '
        '            LogLine = "" _
        '                & LogFileCopyPrep(FormatDateTime(Now(), vbGeneralDate)) _
        '                & "," & LogFileCopyPrep(ContensiveAppName) _
        '                & "," & LogFileCopyPrep(ProgramName) _
        '                & "," & LogFileCopyPrep(ClassName) _
        '                & "," & LogFileCopyPrep(MethodName) _
        '                & "," & LogFileCopyPrep(Context) _
        '                & "," & LogFileCopyPrep(ErrorMessage) _
        '                & "," & LogFileCopyPrep(ResumeMessage) _
        '                & "," & LogFileCopyPrep(ErrSource) _
        '                & "," & LogFileCopyPrep(ErrNumber) _
        '                & "," & LogFileCopyPrep(ErrDescription) _
        '                & "," & LogFileCopyPrep(URL) _
        '                & vbCrLf
        '            '
        '            DayNumber = Day(Now)
        '            MonthNumber = Month(Now)
        '            FilenameNoExt = Year(Now)
        '            If MonthNumber < 10 Then
        '                FilenameNoExt = FilenameNoExt & "0"
        '            End If
        '            FilenameNoExt = FilenameNoExt & MonthNumber
        '            If DayNumber < 10 Then
        '                FilenameNoExt = FilenameNoExt & "0"
        '            End If
        '            FilenameNoExt = LogNamePrefix & FilenameNoExt & DayNumber
        '            If iLogFolder <> "" Then
        '                iLogFolder = iLogFolder & "\"
        '            End If
        '            iLogFolder = GetProgramPath() & "\logs\" & iLogFolder
        '            PathFilenameNoExt = iLogFolder & FilenameNoExt
        '            '
        '            kmafs = CreateObject("kmaFileSystem3.FileSystemClass")
        '            FolderFileList = kmafs.GetFolderFiles2(iLogFolder)
        '            FolderFiles = Split(FolderFileList, vbCrLf)
        '            For Ptr = 0 To UBound(FolderFiles)
        '                If InStr(1, FolderFiles(Ptr), FilenameNoExt & ".log" & ",", vbTextCompare) <> 0 Then
        '                    FileDetails = Split(FolderFiles(Ptr), vbTab)
        '                    fileSize = cp.Utils.EncodeInteger(FileDetails(5))
        '                    Exit For
        '                End If
        '            Next
        '            If fileSize < 10000000 Then
        '                RetryCnt = 0
        '                SaveOK = False
        '                FileSuffix = ""
        '                On Error Resume Next
        '                Do While (Not SaveOK) And (RetryCnt < 10)
        '                    SaveOK = True
        '                    Call kmafs.AppendFile(LCase(PathFilenameNoExt & FileSuffix & ".log"), LogLine)
        '                    If Err.Number <> 0 Then
        '                        If Err.Number = 70 Then
        '                            '
        '                            ' permission denied - happens when more then one process are writing at once, go to the next suffix
        '                            '
        '                            FileSuffix = "-" & CStr(RetryCnt + 1)
        '                            SaveOK = False
        '                        Else
        '                            '
        '                            ' ignore all other errors - this routine logs errors, so there is nothing to do if it fails
        '                            '
        '                        End If
        '                        RetryCnt = RetryCnt + 1
        '                        Err.Clear()
        '                    End If
        '                Loop
        '            End If
        '            kmafs = Nothing
        '            Exit Sub
        '            '
        'ErrorTrap:
        '            Err.Clear()
        '        End Sub
        '        '
        '        '========================================================================
        '        '   HandleError
        '        '       Logs the error and either resumes next, or raises it to the next level
        '        '========================================================================
        '        '
        '        Public Sub HandleError2(ContensiveAppName As String, Context As String, ProgramName As String, ClassName As String, MethodName As String, ErrNumber as integer, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNext As Boolean, URL As String)
        '            '
        '            Call AppendLogFile2(ContensiveAppName, Context, ProgramName, ClassName, MethodName, ErrNumber, ErrSource, ErrDescription, ErrorTrap, ResumeNext, URL, "", "Trace")
        '            '
        '            If Not ResumeNext Then
        '                On Error GoTo 0
        '                If ErrNumber = 0 Then
        '                    Call Err.Raise(KmaErrorInternal, ErrSource, Context)
        '                Else
        '                    Call Err.Raise(ErrNumber, ErrSource, ErrDescription)
        '                End If
        '            End If
        '            '
        '        End Sub
        '        '
        '        '
        '        '
        '        Private Function LogFileCopyPrep(Source) As String
        '            Dim Copy As String
        '            Copy = Source
        '            Copy = Replace(Copy, vbCrLf, " ")
        '            Copy = Replace(Copy, vbLf, " ")
        '            Copy = Replace(Copy, vbCr, " ")
        '            Copy = Replace(Copy, """", """""")
        '            Copy = """" & Copy & """"
        '            LogFileCopyPrep = Copy
        '        End Function
        '        ' moved to csv
        '        ''
        '        ''=================================================================================================================
        '        ''   GetAddonOptionStringValue
        '        ''
        '        ''   gets the value from a list matching the name
        '        ''
        '        ''   InstanceOptionstring is an "AddonEncoded" name=AddonEncodedValue[selector]descriptor&name=value string
        '        ''=================================================================================================================
        '        ''
        '        'Public Function GetAddonOptionStringValue(OptionName As String, AddonOptionString As String) As String
        '        '    On Error GoTo ErrorTrap
        '        '    '
        '        '    Dim Pos as integer
        '        '    Dim s As String
        '        '    '
        '        '    s = GetArgument(OptionName, AddonOptionString, "", "&")
        '        '    Pos = InStr(1, s, "[")
        '        '    If Pos > 0 Then
        '        '        s = Left(s, Pos - 1)
        '        '    End If
        '        '    s = decodeNvaArgument(s)
        '        '    '
        '        '    GetAddonOptionStringValue = Trim(s)
        '        '    '
        '        '    Exit Function
        '        'ErrorTrap:
        '        '    Call HandleError2("", "", App.EXEName, "ccCommonModule", "GetAddonOptionStringValue", Err.Number, Err.Source, Err.Description, True, False, "")
        '        'End Function
        '        '
        '        '
        '        '
        '        Public Function GetIconSprite(TagID As String, SpriteColumn as integer, IconSrc As String, IconWidth as integer, IconHeight as integer, IconAlt As String, IconTitle As String, onDblClick As String, IconIsInline As Boolean, ACInstanceID As String) As String
        '            '
        '            Dim ImgStyle As String
        '            '
        '            GetIconSprite = "<img" _
        '                & " border=0" _
        '                & " id=""" & TagID & """" _
        '                & " onMouseOver=""this.style.backgroundPosition='" & (-1 * SpriteColumn * IconWidth) & "px -" & (2 * IconHeight) & "px';""" _
        '                & " onMouseOut=""this.style.backgroundPosition='" & (-1 * SpriteColumn * IconWidth) & "px 0px'""" _
        '                & " onDblClick=""" & onDblClick & """" _
        '                & " alt=""" & IconAlt & """" _
        '                & " title=""" & IconTitle & """" _
        '                & " src=""/ccLib/images/spacer.gif"""
        '            ImgStyle = "background:url(" & IconSrc & ") " & (-1 * SpriteColumn * IconWidth) & "px 0px no-repeat;"
        '            ImgStyle = ImgStyle & "width:" & IconWidth & "px;"
        '            ImgStyle = ImgStyle & "height:" & IconHeight & "px;"
        '            If IconIsInline Then
        '                ImgStyle = ImgStyle & "vertical-align:middle;display:inline;"
        '            Else
        '                ImgStyle = ImgStyle & "display:block;"
        '            End If
        '            If ACInstanceID <> "" Then
        '                GetIconSprite = GetIconSprite & " ACInstanceID=""" & ACInstanceID & """"
        '            End If
        '            GetIconSprite = GetIconSprite & " style=""" & ImgStyle & """>"
        '        End Function
        '        '
        '        '
        '        '
        '        Public Function RegGetValue$(MainKey&, SubKey$, Value$)
        '            ' MainKey must be one of the Publicly declared HKEY constants.
        '            Dim sKeyType&       'to return the key type.  This function expects REG_SZ or REG_DWORD
        '            Dim ret&            'returned by registry functions, should be 0&
        '            Dim lpHKey&         'return handle to opened key
        '            Dim lpcbData&       'length of data in returned string
        '            Dim ReturnedString$ 'returned string value
        '            Dim ReturnedLong&   'returned long value
        '            If MainKey >= &H80000000 And MainKey <= &H80000006 Then
        '                ' Open key
        '                ret = RegOpenKeyExA(MainKey, SubKey, 0&, KEY_READ, lpHKey)
        '                If ret <> ERROR_SUCCESS Then
        '                    RegGetValue = ""
        '                    Exit Function     'No key open, so leave
        '                End If

        '                ' Set up buffer for data to be returned in.
        '                ' Adjust next value for larger buffers.
        '                lpcbData = 255
        '                ReturnedString = Space$(lpcbData)

        '                ' Read key
        '      ret& = RegQueryValueExA(lpHKey, Value, ByVal 0&, sKeyType, ReturnedString, lpcbData)
        '                If ret <> ERROR_SUCCESS Then
        '                    RegGetValue = ""   'Value probably doesn't exist
        '                Else
        '                    If sKeyType = REG_DWORD Then
        '            ret = RegQueryValueEx(lpHKey, Value, ByVal 0&, sKeyType, ReturnedLong, 4)
        '                        If ret = ERROR_SUCCESS Then RegGetValue = CStr(ReturnedLong)
        '                    Else
        '                        RegGetValue = Left$(ReturnedString, lpcbData - 1)
        '                    End If
        '                End If
        '                ' Always close opened keys.
        '                ret = RegCloseKey(lpHKey)
        '            End If
        '        End Function

    End Module
End Namespace
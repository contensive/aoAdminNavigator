VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetFullClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'========================================================================
'
'========================================================================
'
'Private AdminContentCellBackgroundColor As String
'
'Private Main As ccWeb3.MainClass
Private Csv As Object
Private Main As Object
'
'=================================================================================
'   Execute Method, v3.4 Interface
'=================================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    '
    Set Main = MainObject
    Set Csv = CsvObject
    '
    '
    ' For Page Addons, return the result
    '
    Execute = GetFullNavigator()
    '
    Exit Function
ErrorTrap:
    'HandleError
End Function
'
'=================================================================================
'   Get the Full Navigator, headers and all
'=================================================================================
'
Public Function GetFullNavigator() As String
    On Error GoTo ErrorTrap
    '
    Const NavigatorClosedLabel = "<div style=""font-size:9px;text-align:center;"">&nbsp;<BR>N<BR>a<BR>v<BR>i<BR>g<BR>a<BR>t<BR>o<BR>r</div>"
    '
    Dim OpenNodeList As String
    Dim AdminNavOpen As Boolean
    Dim AdminNavContent As String
    Dim AdminNavHead As String
    Dim AdminNavJS As String
    Dim AdminNavHeadOpened As String
    Dim AdminNavHeadClosed As String
    Dim AdminNavContentOpened As String
    Dim AdminNavContentClosed As String
    Dim AdminNav As String
    Dim GetNode As GetNodeClass
    '
    OpenNodeList = Main.GetVisitProperty("AdminNavOpenNodeList", "")
    AdminNavOpen = kmaEncodeBoolean(Main.GetVisitProperty("AdminNavOpen", "1"))
    If AdminNavOpen Then
        '
        ' draw the page with Nav Opened
        '
        Set GetNode = New GetNodeClass
        AdminNav = GetNode.Execute(Csv, Main, "nodeid=0", "")
        AdminNavHeadOpened = "" _
            & CR & "<div id=""AdminNavHeadOpened"" class=""opened"">" _
            & CR & vbTab & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>" _
            & CR & vbTab & vbTab & "<td valign=Middle class=""left"">Navigator</td>" _
            & CR & vbTab & vbTab & "<td valign=Middle class=""right""><a href=""#"" onClick=""CloseAdminNav();return false""><img alt=""Close Navigator"" title=""Close Navigator"" src=""/cclib/images/ClosexRev1313.gif"" width=13 height=13 border=0></a></td>" _
            & CR & vbTab & "</tr></table>" _
            & CR & "</div>" _
            & ""
        AdminNavHeadClosed = "" _
            & CR & "<div id=""AdminNavHeadClosed"" class=""closed"" style=""display:none;"">" _
            & CR & vbTab & "<a href=""#"" onClick=""OpenAdminNav();return false""><img title=""Open Navigator"" alt=""Open Navigator"" src=""/cclib/images/OpenRightRev1313.gif"" width=13 height=13 border=0 style=""text-align:right;""></a>" _
            & CR & "</div>" _
            & ""
        AdminNavHead = "" _
            & CR & "<div class=""ccHeaderCon"">" _
            & KmaIndent(AdminNavHeadOpened) _
            & KmaIndent(AdminNavHeadClosed) _
            & CR & "</div>" _
            & ""
        AdminNavContentOpened = "" _
            & CR & "<div id=""AdminNavContentOpened"" class=""opened"">" _
            & KmaIndent(AdminNav & "<img alt=""space"" src=""/cclib/images/spacer.gif"" width=""200"" height=""1"" style=""clear:both"">") _
            & CR & "</div>" _
            & ""
        AdminNavContentClosed = "" _
            & CR & "<div id=""AdminNavContentClosed"" class=""closed"" style=""display:none;"">" & NavigatorClosedLabel & "</div>" _
            & ""
        AdminNavContent = "" _
            & CR & "<div class=""ccContentCon"">" _
            & KmaIndent(AdminNavContentOpened) _
            & KmaIndent(AdminNavContentClosed) _
            & CR & "</div>" _
            & ""
'                & CR & "<div class=""ccContentCon"">" _
'                & CR & "<div id=""AdminNavContentOpened"" class=""opened"">" _
'                & KmaIndent(GetNavigator(Main, "0", OpenNodeList) & "<img alt=""space"" src=""/cclib/images/spacer.gif"" width=""200"" height=""1"" style=""clear:both"">") _
'                & CR & "</div>" _
'                & CR & "<div id=""AdminNavContentClosed"" class=""closed"" style=""display:none;"">" & NavigatorClosedLabel & "</div>" _
'                & CR & "</div>"
        AdminNavJS = "" _
            & CR & "function CloseAdminNav() {SetDisplay('AdminNavHeadOpened','none');SetDisplay('AdminNavContentOpened','none');SetDisplay('AdminNavHeadClosed','block');SetDisplay('AdminNavContentClosed','block');cj.ajax.setVisitProperty('','AdminNavOpen','0')}" _
            & CR & "function OpenAdminNav() {SetDisplay('AdminNavHeadOpened','block');SetDisplay('AdminNavContentOpened','block');SetDisplay('AdminNavHeadClosed','none');SetDisplay('AdminNavContentClosed','none');cj.ajax.setVisitProperty('','AdminNavOpen','1')}" _
            & CR & ""
'        AdminNavJS = "" _
'            & CR & "function CloseAdminNav() {SetDisplay('AdminNavHeadOpened','none');SetDisplay('AdminNavContentOpened','none');SetDisplay('AdminNavHeadClosed','block');SetDisplay('AdminNavContentClosed','block');cj.ajax.qs('" & RequestNameAjaxFunction & "=" & AjaxCloseAdminNav & "','','')}" _
'            & CR & "function OpenAdminNav() {SetDisplay('AdminNavHeadOpened','block');SetDisplay('AdminNavContentOpened','block');SetDisplay('AdminNavHeadClosed','none');SetDisplay('AdminNavContentClosed','none');cj.ajax.qs('" & RequestNameAjaxFunction & "=" & AjaxOpenAdminNav & "','','')}" _
'            & CR & ""
    Else
        '
        ' draw the page with Nav Closed
        '
        AdminNavHeadOpened = "" _
            & CR & "<div id=""AdminNavHeadOpened"" class=""opened"" style=""display:none;"">" _
            & CR & vbTab & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>" _
            & CR & vbTab & vbTab & "<td valign=Middle class=""left"">Navigator</td>" _
            & CR & vbTab & vbTab & "<td valign=Middle class=""right""><a href=""#"" onClick=""CloseAdminNav();return false""><img alt=""Close Navigator"" title=""Close Navigator"" src=""/cclib/images/ClosexRev1313.gif"" width=13 height=13 border=0></a></td>" _
            & CR & vbTab & "</tr></table>" _
            & CR & "</div>" _
            & ""
        AdminNavHeadClosed = "" _
            & CR & "<div id=""AdminNavHeadClosed"" class=""closed"">" _
            & CR & vbTab & "<a href=""#"" onClick=""OpenAdminNav();return false""><img title=""Open Navigator"" alt=""Open Navigator"" src=""/cclib/images/OpenRightRev1313.gif"" width=13 height=13 border=0 style=""text-align:right;""></a>" _
            & CR & "</div>" _
            & ""
        AdminNavHead = "" _
            & CR & "<div class=""ccHeaderCon"">" _
            & KmaIndent(AdminNavHeadOpened) _
            & KmaIndent(AdminNavHeadClosed) _
            & CR & "</div>" _
            & ""
        AdminNavContent = "" _
            & CR & "<div class=""ccContentCon"">" _
            & CR & "<div id=""AdminNavContentOpened"" class=""opened"" style=""display:none;""><div style=""text-align:center;""><img src=""/cclib/images/ajax-loader-small.gif"" width=16 height=16></div></div>" _
            & CR & "<div id=""AdminNavContentClosed"" class=""closed"">" & NavigatorClosedLabel & "</div>" _
            & CR & "<div id=""AdminNavContentMinWidth"" style=""display:none;""><img alt=""space"" src=""/cclib/images/spacer.gif"" width=""200"" height=""1"" style=""clear:both""></div>" _
            & CR & "</div>"
        AdminNavJS = "" _
            & CR & "var AdminNavPop=false;" _
            & CR & "function CloseAdminNav() {SetDisplay('AdminNavHeadOpened','none');SetDisplay('AdminNavHeadClosed','block');SetDisplay('AdminNavContentOpened','none');SetDisplay('AdminNavContentMinWidth','none');SetDisplay('AdminNavContentClosed','block');cj.ajax.setVisitProperty('','AdminNavOpen','0')}" _
            & CR & "function OpenAdminNav() {SetDisplay('AdminNavHeadOpened','block');SetDisplay('AdminNavHeadClosed','none');SetDisplay('AdminNavContentOpened','block');SetDisplay('AdminNavContentMinWidth','block');SetDisplay('AdminNavContentClosed','none');cj.ajax.setVisitProperty('','AdminNavOpen','1');if(!AdminNavPop){cj.ajax.addon('AdminNavigatorGetNode','','','AdminNavContentOpened');AdminNavPop=true;}else{cj.ajax.addon('AdminNavigatorOpenNode');}}" _
            & CR & ""
    End If
    Call Main.AddHeadScriptCode(AdminNavJS, "Admin Navigator")
    '
    '
    '
    GetFullNavigator = "" _
        & CR & "<div>" _
        & KmaIndent(AdminNavHead) _
        & KmaIndent(AdminNavContent) _
        & CR & "</div>" _
        & ""
    '
    Exit Function
ErrorTrap:
    'HandleError
End Function

'
'
'
Private Sub HandleClassAppendLogfile(MethodName As String, Context As String)
    Call AppendLogFile2(Main.ApplicationName, Context, "ccWeb3", "AddonManClass", MethodName, 0, "", "", False, True, Main.ServerLink, "", "trace")

End Sub
'
'===========================================================================
'
'===========================================================================
'
Private Sub HandleClassTrapError(MethodName As String, Optional Context As String)
    '
    If Main Is Nothing Then
        Call HandleError2("unknown", Context, App.EXEName, "AddonManClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, "unknown")
    Else
        Call HandleError2(Main.ApplicationName, Context, App.EXEName, "AddonManClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, Main.ServerLink)
    End If
    '
    End Sub


Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.adminNavigator
    '
    ' Sample Vb addon
    '
    Public Class getFullClass
        Inherits AddonBaseClass
        '
        ' - Change this class name to the addon name
        ' - Create a Contensive Addon record, set the dotnet class full name to yourNameSpaceName.yourClassName
        '
        '=====================================================================================
        ' addon api
        '=====================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                returnHtml = GetFullNavigator(CP)
            Catch ex As Exception
                errorReport(CP, ex, "execute")
            End Try
            Return DirectCast(returnHtml, String)
        End Function
        '
        '=================================================================================
        '   Get the Full Navigator, headers and all
        '=================================================================================
        '
        Public Function GetFullNavigator(cp As CPBaseClass) As String
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
            Dim GetNode As getNodeClass
            '
            OpenNodeList = cp.Visit.GetText("AdminNavOpenNodeList", "")
            AdminNavOpen = cp.Utils.EncodeBoolean(cp.Visit.GetText("AdminNavOpen", "1"))
            If AdminNavOpen Then
                '
                ' draw the page with Nav Opened
                '
                GetNode = New getNodeClass
                cp.Doc.SetProperty("nodeid", "0")
                AdminNav = GetNode.Execute(cp).ToString
                AdminNavHeadOpened = "" _
                    & cr & "<div id=""AdminNavHeadOpened"" class=""opened"">" _
                    & cr & vbTab & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>" _
                    & cr & vbTab & vbTab & "<td valign=Middle class=""left"">Navigator</td>" _
                    & cr & vbTab & vbTab & "<td valign=Middle class=""right""><a href=""#"" onClick=""CloseAdminNav();return false""><img alt=""Close Navigator"" title=""Close Navigator"" src=""/cclib/images/ClosexRev1313.gif"" width=13 height=13 border=0></a></td>" _
                    & cr & vbTab & "</tr></table>" _
                    & cr & "</div>" _
                    & ""
                AdminNavHeadClosed = "" _
                    & cr & "<div id=""AdminNavHeadClosed"" class=""closed"" style=""display:none;"">" _
                    & cr & vbTab & "<a href=""#"" onClick=""OpenAdminNav();return false""><img title=""Open Navigator"" alt=""Open Navigator"" src=""/cclib/images/OpenRightRev1313.gif"" width=13 height=13 border=0 style=""text-align:right;""></a>" _
                    & cr & "</div>" _
                    & ""
                AdminNavHead = "" _
                    & cr & "<div class=""ccHeaderCon"">" _
                    & kmaIndent(AdminNavHeadOpened) _
                    & kmaIndent(AdminNavHeadClosed) _
                    & cr & "</div>" _
                    & ""
                AdminNavContentOpened = "" _
                    & cr & "<div id=""AdminNavContentOpened"" class=""opened"">" _
                    & kmaIndent(AdminNav & "<img alt=""space"" src=""/cclib/images/spacer.gif"" width=""200"" height=""1"" style=""clear:both"">") _
                    & cr & "</div>" _
                    & ""
                AdminNavContentClosed = "" _
                    & cr & "<div id=""AdminNavContentClosed"" class=""closed"" style=""display:none;"">" & NavigatorClosedLabel & "</div>" _
                    & ""
                AdminNavContent = "" _
                    & cr & "<div class=""ccContentCon"">" _
                    & kmaIndent(AdminNavContentOpened) _
                    & kmaIndent(AdminNavContentClosed) _
                    & cr & "</div>" _
                    & ""
                '                & CR & "<div class=""ccContentCon"">" _
                '                & CR & "<div id=""AdminNavContentOpened"" class=""opened"">" _
                '                & KmaIndent(GetNavigator(Main, "0", OpenNodeList) & "<img alt=""space"" src=""/cclib/images/spacer.gif"" width=""200"" height=""1"" style=""clear:both"">") _
                '                & CR & "</div>" _
                '                & CR & "<div id=""AdminNavContentClosed"" class=""closed"" style=""display:none;"">" & NavigatorClosedLabel & "</div>" _
                '                & CR & "</div>"
                AdminNavJS = "" _
                    & cr & "function CloseAdminNav() {SetDisplay('AdminNavHeadOpened','none');SetDisplay('AdminNavContentOpened','none');SetDisplay('AdminNavHeadClosed','block');SetDisplay('AdminNavContentClosed','block');cj.ajax.setVisitProperty('','AdminNavOpen','0')}" _
                    & cr & "function OpenAdminNav() {SetDisplay('AdminNavHeadOpened','block');SetDisplay('AdminNavContentOpened','block');SetDisplay('AdminNavHeadClosed','none');SetDisplay('AdminNavContentClosed','none');cj.ajax.setVisitProperty('','AdminNavOpen','1')}" _
                    & cr & ""
                '        AdminNavJS = "" _
                '            & CR & "function CloseAdminNav() {SetDisplay('AdminNavHeadOpened','none');SetDisplay('AdminNavContentOpened','none');SetDisplay('AdminNavHeadClosed','block');SetDisplay('AdminNavContentClosed','block');cj.ajax.qs('" & RequestNameAjaxFunction & "=" & AjaxCloseAdminNav & "','','')}" _
                '            & CR & "function OpenAdminNav() {SetDisplay('AdminNavHeadOpened','block');SetDisplay('AdminNavContentOpened','block');SetDisplay('AdminNavHeadClosed','none');SetDisplay('AdminNavContentClosed','none');cj.ajax.qs('" & RequestNameAjaxFunction & "=" & AjaxOpenAdminNav & "','','')}" _
                '            & CR & ""
            Else
                '
                ' draw the page with Nav Closed
                '
                AdminNavHeadOpened = "" _
                    & cr & "<div id=""AdminNavHeadOpened"" class=""opened"" style=""display:none;"">" _
                    & cr & vbTab & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>" _
                    & cr & vbTab & vbTab & "<td valign=Middle class=""left"">Navigator</td>" _
                    & cr & vbTab & vbTab & "<td valign=Middle class=""right""><a href=""#"" onClick=""CloseAdminNav();return false""><img alt=""Close Navigator"" title=""Close Navigator"" src=""/cclib/images/ClosexRev1313.gif"" width=13 height=13 border=0></a></td>" _
                    & cr & vbTab & "</tr></table>" _
                    & cr & "</div>" _
                    & ""
                AdminNavHeadClosed = "" _
                    & cr & "<div id=""AdminNavHeadClosed"" class=""closed"">" _
                    & cr & vbTab & "<a href=""#"" onClick=""OpenAdminNav();return false""><img title=""Open Navigator"" alt=""Open Navigator"" src=""/cclib/images/OpenRightRev1313.gif"" width=13 height=13 border=0 style=""text-align:right;""></a>" _
                    & cr & "</div>" _
                    & ""
                AdminNavHead = "" _
                    & cr & "<div class=""ccHeaderCon"">" _
                    & kmaIndent(AdminNavHeadOpened) _
                    & kmaIndent(AdminNavHeadClosed) _
                    & cr & "</div>" _
                    & ""
                AdminNavContent = "" _
                    & cr & "<div class=""ccContentCon"">" _
                    & cr & "<div id=""AdminNavContentOpened"" class=""opened"" style=""display:none;""><div style=""text-align:center;""><img src=""/cclib/images/ajax-loader-small.gif"" width=16 height=16></div></div>" _
                    & cr & "<div id=""AdminNavContentClosed"" class=""closed"">" & NavigatorClosedLabel & "</div>" _
                    & cr & "<div id=""AdminNavContentMinWidth"" style=""display:none;""><img alt=""space"" src=""/cclib/images/spacer.gif"" width=""200"" height=""1"" style=""clear:both""></div>" _
                    & cr & "</div>"
                AdminNavJS = "" _
                    & cr & "var AdminNavPop=false;" _
                    & cr & "function CloseAdminNav() {SetDisplay('AdminNavHeadOpened','none');SetDisplay('AdminNavHeadClosed','block');SetDisplay('AdminNavContentOpened','none');SetDisplay('AdminNavContentMinWidth','none');SetDisplay('AdminNavContentClosed','block');cj.ajax.setVisitProperty('','AdminNavOpen','0')}" _
                    & cr & "function OpenAdminNav() {SetDisplay('AdminNavHeadOpened','block');SetDisplay('AdminNavHeadClosed','none');SetDisplay('AdminNavContentOpened','block');SetDisplay('AdminNavContentMinWidth','block');SetDisplay('AdminNavContentClosed','none');cj.ajax.setVisitProperty('','AdminNavOpen','1');if(!AdminNavPop){cj.ajax.addon('AdminNavigatorGetNode','','','AdminNavContentOpened');AdminNavPop=true;}else{cj.ajax.addon('AdminNavigatorOpenNode');}}" _
                    & cr & ""
            End If
            Callcp.Doc.AddHeadJavascript(AdminNavJS, "Admin Navigator")
            '
            '
            '
            GetFullNavigator = "" _
                & cr & "<div>" _
                & kmaIndent(AdminNavHead) _
                & kmaIndent(AdminNavContent) _
                & cr & "</div>" _
                & ""
            '
            Exit Function
ErrorTrap:
            'HandleError
        End Function

        '
        '
        '
        '
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in sampleClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
    End Class
End Namespace

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    Public Class GetFullClass
        Inherits AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Try
                Dim returnHtml As String = ""
                Dim AdminNavContent As String
                Dim AdminNavHead As String
                Dim AdminNavHeadOpened As String
                Dim AdminNavHeadClosed As String
                Dim env As New NavigatorEnvironment(CP)
                If env.AdminNavOpen Then
                    '
                    ' draw the page with Nav Opened
                    '
                    CP.Doc.SetProperty("nodeid", "0")
                    AdminNavHeadOpened = "" _
                        & cr & "<div id=""AdminNavHeadOpened"" class=""opened"">" _
                        & cr & vbTab & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>" _
                        & cr & vbTab & vbTab & "<td valign=Middle class=""left"">Navigator</td>" _
                        & cr & vbTab & vbTab & "<td valign=Middle class=""right""><a href=""#"" onClick=""closeAdminNav();return false"">" & iconCloseWindow & "</a></td>" _
                        & cr & vbTab & "</tr></table>" _
                        & cr & "</div>" _
                        & ""
                    AdminNavHeadClosed = "" _
                        & cr & "<div id=""AdminNavHeadClosed"" class=""closed"" style=""display:none;"">" _
                        & cr & vbTab & "<a href=""#"" onClick=""reOpenAdminNav();return false"">" & iconOpenWindow & "</a>" _
                        & cr & "</div>" _
                        & ""
                    AdminNavHead = "" _
                        & cr & "<div class=""ccHeaderCon"">" _
                        & kmaIndent(AdminNavHeadOpened) _
                        & kmaIndent(AdminNavHeadClosed) _
                        & cr & "</div>" _
                        & ""
                    Dim AdminNavContentOpened As String = "" _
                        & cr & "<div id=""AdminNavContentOpened"" class=""opened"">" _
                        & kmaIndent(GetNodeClass.getNode(CP, env).ToString & "<img alt=""space"" src=""https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/spacer.gif"" width=""200"" height=""1"" style=""clear:both"">") _
                        & cr & "</div>" _
                        & ""
                    Dim AdminNavContentClosed As String = "" _
                        & cr & "<div id=""AdminNavContentClosed"" class=""closed"" style=""display:none;"">" & NavigatorClosedLabel & "</div>" _
                        & ""
                    AdminNavContent = "" _
                        & cr & "<div class=""ccContentCon"">" _
                        & kmaIndent(AdminNavContentOpened) _
                        & kmaIndent(AdminNavContentClosed) _
                        & cr & "</div>" _
                        & ""
                Else
                    '
                    ' draw the page with Nav Closed
                    '
                    AdminNavHeadOpened = "" _
                        & cr & "<div id=""AdminNavHeadOpened"" class=""opened"" style=""display:none;"">" _
                        & cr & vbTab & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>" _
                        & cr & vbTab & vbTab & "<td valign=Middle class=""left"">Navigator</td>" _
                        & cr & vbTab & vbTab & "<td valign=Middle class=""right""><a href=""#"" onClick=""reCloseAdminNav();return false"">" & iconCloseWindow & "</a></td>" _
                        & cr & vbTab & "</tr></table>" _
                        & cr & "</div>" _
                        & ""
                    AdminNavHeadClosed = "" _
                        & cr & "<div id=""AdminNavHeadClosed"" class=""closed"">" _
                        & cr & vbTab & "<a href=""#"" onClick=""OpenAdminNav();return false"">" & iconOpenWindow & "</a>" _
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
                        & cr & "<div id=""AdminNavContentOpened"" class=""opened"" style=""display:none;""><div style=""text-align:center;""><img src=""https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/ajax-loader-small.gif"" width=16 height=16></div></div>" _
                        & cr & "<div id=""AdminNavContentClosed"" class=""closed"">" & NavigatorClosedLabel & "</div>" _
                        & cr & "<div id=""AdminNavContentMinWidth"" style=""display:none;""><img alt=""space"" src=""https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/spacer.gif"" width=""200"" height=""1"" style=""clear:both""></div>" _
                        & cr & "</div>"
                End If
                returnHtml = "" _
                    & cr & "<div>" _
                    & kmaIndent(AdminNavHead) _
                    & kmaIndent(AdminNavContent) _
                    & cr & "</div>" _
                    & ""
                Return returnHtml
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
    End Class
End Namespace

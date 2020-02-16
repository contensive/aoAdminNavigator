
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions
Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    Public Class NodeController
        '
        '====================================================================================================
        '
        Friend Shared Function GetNode(cp As CPBaseClass,
                                 env As NavigatorEnvironment,
                                 CollectionID As Integer,
                                 ContentControlID As Integer,
                                 helpCollectionID As Integer,
                                 HelpAddonID As Integer,
                                 ContentID As Integer,
                                 Link As String,
                                 addon As Models.Db.AddonModel,
                                 ignore As Integer,
                                 Name As String,
                                 xEmptyNodeList As String,
                                 NavigatorID As Integer,
                                 NavIconType As Integer,
                                 NavIconTitleHtmlEncoded As String,
                                 AutoManageAddons As Boolean,
                                 NodeType As NodeTypeEnum,
                                 NewWindow As Boolean,
                                 BlockSubNodes As Boolean,
                                 xOpenNodeList As String,
                                 NodeIDString As String,
                                 ByRef Return_NavigatorJS As String,
                                 linkSuffixList As String
                                 ) As String
            On Error GoTo ErrorTrap
            '
            Dim csCollection As Integer
            Dim collectionName As String
            Dim collectionHelpLink As String
            Dim collectionHelp As String
            Dim NodeNavigatorJS As String = ""
            Dim SubNav As String
            Dim DivIDClosed As String
            Dim DivIDOpened As String
            Dim DivIDContent As String
            Dim DivIDEmpty As String
            Dim s As String
            Dim CSAddon As Integer
            Dim DivIDBase As String
            Dim IconNoSubNodes As String
            Dim IconOpened As String
            Dim IconClosed As String
            Dim AddonGuid As String
            Dim AddonName As String
            Dim IsVisible As Boolean
            Dim WorkingName As String
            Dim workingNameHtmlEncoded As String
            Dim BlockNode As Boolean
            '
            ' Determine if it has either a function, or child entries
            '
            BlockNode = False
            Return_NavigatorJS = ""
            WorkingName = Name
            IsVisible = (CollectionID <> 0) Or (helpCollectionID <> 0) Or (HelpAddonID <> 0) Or (ContentID <> 0) Or (Link <> "") Or (addon IsNot Nothing) Or (LCase(WorkingName) = "all content") Or (LCase(WorkingName) = "add-ons with no collection")
            If Not IsVisible Then
                '
                ' IsVisible if it is not in the EmptyNodeList (has child entries)
                '
                IsVisible = (InStr(1, "," & env.EmptyNodeList & ",", "," & NavigatorID & ",") = 0)
            End If
            If IsVisible Then
                '
                ' hide the legacy node 'switch to navigator'
                '
                IsVisible = LCase(WorkingName) <> "switch to navigator"
            End If
            If IsVisible Then
                '
                ' Setup Icons
                '
                Select Case NavIconType
                    Case NavIconTypeCustom
                        '
                        ' reserved for future addition of a custom Icon field
                        ' not done now because there is no facility now to import files during collection build
                        '
                    Case NavIconTypeAdvanced
                        IconOpened = IconAdvancedOpened
                        IconClosed = IconAdvancedClosed
                        IconNoSubNodes = IconAdvanced
                    Case NavIconTypeContent
                        IconOpened = IconContentOpened
                        IconClosed = IconContentClosed
                        IconNoSubNodes = IconContent
                    Case NavIconTypeEmail
                        IconOpened = IconEmailOpened
                        IconClosed = IconEmailClosed
                        IconNoSubNodes = IconEmail
                    Case NavIconTypeUser
                        IconOpened = IconUsersOpened
                        IconClosed = IconUsersClosed
                        IconNoSubNodes = IconUsers
                    Case NavIconTypeReport
                        IconOpened = IconReportsOpened
                        IconClosed = IconReportsClosed
                        IconNoSubNodes = IconReports
                    Case NavIconTypeSetting
                        IconOpened = IconSettingsOpened
                        IconClosed = IconSettingsClosed
                        IconNoSubNodes = IconSettings
                    Case NavIconTypeTool
                        IconOpened = IconToolsOpened
                        IconClosed = IconToolsClosed
                        IconNoSubNodes = IconTools
                    Case NavIconTypeRecord
                        IconOpened = IconRecordOpened
                        IconClosed = IconRecordClosed
                        IconNoSubNodes = IconRecord
                    Case NavIconTypeAddon
                        IconOpened = IconAddonsOpened
                        IconClosed = IconAddonsClosed
                        IconNoSubNodes = IconAddons
                    Case NavIconTypeHelp
                        IconOpened = IconHelp
                        IconClosed = IconHelp
                        IconNoSubNodes = IconHelp
                    Case Else 'NavIconTypeFolder
                        IconOpened = IconFolderOpened
                        IconClosed = IconFolderClosed
                        IconNoSubNodes = IconFolderNoSubNodes
                End Select
                IconOpened = Replace(IconOpened, "{title}", "Close " & NavIconTitleHtmlEncoded)
                IconClosed = Replace(IconClosed, "{title}", "Open " & NavIconTitleHtmlEncoded)
                IconNoSubNodes = Replace(IconNoSubNodes, "{title}", NavIconTitleHtmlEncoded)
                '
                ' NodeIDString - the unique string that is passed by here as ParentNode to get all the child nodes
                '   is always the navigator entry ID, unless it is a hardcoded subsection
                ' DIVID must be unique for this entire menu, but does not need to be recognized in a call back to the server
                '
                '
                ' set flag for 'hardcoded' lists - like add-ons
                '
                If AutoManageAddons And (LCase(WorkingName) = "manage add-ons") Then
                    '
                    ' test special case - replace manage add-ons branch
                    '
                    DivIDBase = CStr(NavigatorID)
                    'NodeIDString = NodeIDManageAddons
                Else
                    Select Case LCase(WorkingName)
                        Case "legacy menu"
                            '
                            ' special case - if node has this name, a click to it calls back with NodeIDLegacyMenu
                            '

                            DivIDBase = CStr(NavigatorID)
                            'NodeIDString = NodeIDLegacyMenu
                        Case "add-ons with no collection"
                            '
                            ' any Navigator Entry named 'all content' returns the content list
                            '
                            DivIDBase = CStr(NavigatorID)
                            'NodeIDString = NodeIDAddonsNoCollection
                        Case "all content"
                            '
                            ' any Navigator Entry named 'all content' returns the content list
                            '
                            DivIDBase = CStr(NavigatorID)
                            'NodeIDString = NodeIDAllContentList
                        Case Else
                            '
                            ' This entry is made from a navigator entry record
                            '
                            Select Case NodeType
                                Case NodeTypeEnum.NodeTypeAddon
                                    '
                                    ' List of Addons
                                    '
                                    'NodeIDString = ""
                                    DivIDBase = "a" & NavigatorID
                                Case NodeTypeEnum.NodeTypeCollection
                                    '
                                    ' List of Collections
                                    '
                                    'NodeIDString = "collection." & CollectionID
                                    DivIDBase = "c" & CollectionID
                                Case NodeTypeEnum.NodeTypeContent
                                    '
                                    ' List of content
                                    '
                                    'NodeIDString = ""
                                    DivIDBase = "d" & NavigatorID
                                Case Else
                                    '
                                    ' List of Navigator Entries
                                    '
                                    'NodeIDString = CStr(NavigatorID)
                                    DivIDBase = "n" & NavigatorID
                            End Select
                    End Select
                End If
                '
                ' check name for length
                '
                If Len(WorkingName) > 53 Then
                    WorkingName = Left(WorkingName, 25) & "..." & Right(WorkingName, 25)
                End If
                Dim NavLinkHTMLId As String = ""
                '
                ' setup link
                '
                If Link <> "" Then
                    NavLinkHTMLId = "n" & NavigatorID
                    workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                    If NewWindow Then
                        workingNameHtmlEncoded = "<a class=navDrag name=""navLink"" id=""" & NavLinkHTMLId & """ href=""" & Link & """ target=""_blank"" title=""Open '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                    Else
                        workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & Link & """ title=""Open '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                    End If
                Else
                    '
                    ' If link page, use this
                    '
                    If (addon IsNot Nothing) Then
                        '
                        ' link to addon
                        '
                        Link = ""
                        AddonGuid = addon.ccguid
                        AddonName = addon.name
                        If addon.remoteMethod Then
                            NewWindow = True
                            Link = cp.Site.GetText("adminUrl") & "?" & RequestNameRemoteMethodAddon & "=" & cp.Utils.EncodeRequestVariable(AddonName)
                        End If
                        If Link = "" Then
                            If AddonGuid <> "" Then
                                Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "addonguid", AddonGuid, True)
                            Else
                                Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "addonid", CStr(addon.id), True)
                            End If
                        End If
                        NavLinkHTMLId = "a" & addon.id
                        workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                        If NewWindow Then
                            workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & Link & """ target=""_blank"" title=""Run '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                        Else
                            workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & Link & """ title=""Run '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                        End If
                        '
                        If (env.allowToolTip And (addon.template Or addon.content Or (addon.navTypeId > 1))) Then
                            '
                            ' -- add tooltip for addons
                            workingNameHtmlEncoded &= "<i data-tooltip=""" & AddonGuid & """ class=""contensiveToolTip fas fa-info-circle"" href=""#"" data-toggle=""tooltip"" data-html=""true"" rel=""tooltip"" title=""Click for details.""></i>"
                        End If
                    ElseIf ContentID <> 0 Then
                        '
                        ' go edit the content
                        '
                        Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "cid", CStr(ContentID), True)
                        NavLinkHTMLId = "c" & ContentID
                        workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                        If NewWindow Then
                            workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & Link & """ target=""_blank"" title=""List All '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                        Else
                            workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & Link & """ title=""List All '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                        End If
                    ElseIf HelpAddonID <> 0 Then
                        '
                        ' go to Addon Help
                        '
                        Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "helpaddonid", CStr(HelpAddonID), True)
                        workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                        If NewWindow Then
                            workingNameHtmlEncoded = "<a href=""" & Link & """ target=""_blank"" title=""Help for Add-on '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                        Else
                            workingNameHtmlEncoded = "<a href=""" & Link & """ title=""Help for Add-on '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                        End If
                    ElseIf helpCollectionID <> 0 Then
                        '
                        ' go to Collection Help
                        '
                        Using cs13 As CPCSBaseClass = cp.CSNew
                            BlockNode = True
                            If cs13.Open("add-on collections", "id=" & helpCollectionID, "name", True, "name,helpLink,help") Then

                                collectionName = cs13.GetText("name")
                                collectionHelpLink = cs13.GetText("helpLink")
                                collectionHelp = cs13.GetText("help")
                                '
                                WorkingName = collectionName
                                Link = collectionHelpLink
                                workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                                If Link <> "" Then
                                    BlockNode = False
                                    NewWindow = True
                                ElseIf (collectionHelp <> "") Then
                                    BlockNode = False
                                    Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "helpcollectionid", CStr(helpCollectionID), True)
                                End If
                                If Not BlockNode Then
                                    If NewWindow Then
                                        workingNameHtmlEncoded = "<a href=""" & Link & """ target=""_blank"" title=""Help for Collection '" & workingNameHtmlEncoded & "'"">Help</a>"
                                    Else
                                        workingNameHtmlEncoded = "<a href=""" & Link & """ title=""Help for Collection '" & workingNameHtmlEncoded & "'"">Help</a>"
                                    End If
                                End If

                            End If
                            cs13.Close()
                        End Using
                    Else
                        workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                    End If
                End If
                '
                If Not BlockNode Then
                    If BlockSubNodes Or (addon IsNot Nothing) Then
                        '
                        ' This is a hardcoded item (like Add-on), it has no subnodes
                        '
                        s &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                    Else
                        '
                        DivIDClosed = DivIDBase & "a"
                        DivIDOpened = DivIDBase & "b"
                        DivIDContent = DivIDBase & "c"
                        DivIDEmpty = DivIDBase & "d"

                        If (ContentID = 0) And (InStr(1, env.EmptyNodeList & ",", "," & NodeIDString & ",") <> 0) Then
                            '
                            ' In EmptyNodeList
                            '
                            s &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                        ElseIf InStr(1, env.OpenNodeList & ",", "," & NodeIDString & ",") <> 0 Then
                            '
                            ' This node is open
                            '
                            SubNav = NodeListController.getNodeList(cp, env, NodeIDString, NodeNavigatorJS)
                            Return_NavigatorJS &= NodeNavigatorJS
                            If SubNav <> "" Then
                                '
                                ' display the subnav
                                '
                                s &= cr & "<div class=ccNavLink ID=" & DivIDClosed & " style=""display:none;""><A class=""ccNavClosed"" href=""#"" onclick=""AdminNavOpenClick('" & DivIDClosed & "','" & DivIDOpened & "','" & DivIDContent & "','" & NodeIDString & "','" & DivIDEmpty & "');return false;"">" & IconClosed & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                s &= cr & "<div class=ccNavLink ID=" & DivIDOpened & "><A class=""ccNavOpened"" href=""#"" onclick=""AdminNavCloseClick('" & DivIDOpened & "','" & DivIDClosed & "','" & DivIDContent & "','" & NodeIDString & "');return false;"">" & IconOpened & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                s = s _
                                    & cr & "<div class=ccNavLinkChild ID=" & DivIDContent & ">" _
                                    & kmaIndent(SubNav) _
                                    & cr & "</div>"
                            Else
                                '
                                ' it has a NO subnav
                                '
                                s &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            End If
                        Else
                            '
                            ' This node is closed
                            '
                            s &= cr & "<div class=ccNavLink ID=" & DivIDClosed & " ><A class=""ccNavClosed"" href=""#"" onclick=""AdminNavOpenClick('" & DivIDClosed & "','" & DivIDOpened & "','" & DivIDContent & "','" & NodeIDString & "','','" & DivIDContent & "');return false;"">" & IconClosed & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            s &= cr & "<div class=ccNavLink ID=" & DivIDOpened & " style=""display:none;""><A class=""ccNavOpened"" href=""#"" onclick=""AdminNavCloseClick('" & DivIDOpened & "','" & DivIDClosed & "','" & DivIDContent & "','" & NodeIDString & "');return false;"">" & IconOpened & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            s &= cr & "<div class=""ccNavLink ccNavLinkEmpty"" ID=" & DivIDEmpty & " style=""display:none;"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            s &= cr & "<div class=ccNavLinkChild ID=" & DivIDContent & " style=""display:none;margin-left:20px;"">&nbsp;&nbsp;&nbsp;&nbsp;<img src=""https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/ajax-loader-small.gif"" width=""16"" height=""16""></div>"
                        End If
                    End If
                End If
            End If
            '
            GetNode = s
            '
            Exit Function
            '
            ' ----- Error Trap
            '
ErrorTrap:
            Call cp.Site.ErrorReport("Trap")
        End Function
        '
    End Class
End Namespace


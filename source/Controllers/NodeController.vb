
Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    Public NotInheritable Class NodeController
        '
        '====================================================================================================
        ''' <summary>
        ''' get a single node and its children
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="env"></param>
        ''' <param name="CollectionID"></param>
        ''' <param name="ContentControlID"></param>
        ''' <param name="helpCollectionID"></param>
        ''' <param name="HelpAddonID"></param>
        ''' <param name="ContentID"></param>
        ''' <param name="link"></param>
        ''' <param name="addon"></param>
        ''' <param name="ignore"></param>
        ''' <param name="Name"></param>
        ''' <param name="xEmptyNodeList"></param>
        ''' <param name="NavigatorID"></param>
        ''' <param name="NavIconType"></param>
        ''' <param name="NavIconTitleHtmlEncoded"></param>
        ''' <param name="AutoManageAddons"></param>
        ''' <param name="NodeType"></param>
        ''' <param name="NewWindow"></param>
        ''' <param name="BlockSubNodes"></param>
        ''' <param name="xOpenNodeList"></param>
        ''' <param name="NodeIDString"></param>
        ''' <param name="Return_NavigatorJS"></param>
        ''' <param name="linkSuffixList"></param>
        ''' <returns></returns>
        Friend Shared Function getNode(cp As CPBaseClass,
                                 env As ApplicationEnvironmentModel,
                                 CollectionID As Integer,
                                 ContentControlID As Integer,
                                 helpCollectionID As Integer,
                                 HelpAddonID As Integer,
                                 ContentID As Integer,
                                 link As String,
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
            Dim hint As Integer = 0
            Try
                Dim SubNav As String
                Dim DivIDClosed As String
                Dim DivIDOpened As String
                Dim DivIDContent As String
                Dim DivIDEmpty As String
                Dim result As String = ""
                Dim DivIDBase As String
                Dim AddonGuid As String
                Dim AddonName As String
                Dim IsVisible As Boolean
                Dim WorkingName As String
                '
                ' Determine if it has either a function, or child entries
                '
                Dim BlockNode As Boolean = False
                Return_NavigatorJS = ""
                WorkingName = Name
                IsVisible = (CollectionID <> 0) Or (helpCollectionID <> 0) Or (HelpAddonID <> 0) Or (ContentID <> 0) Or Not String.IsNullOrEmpty(link) Or (addon IsNot Nothing) Or (LCase(WorkingName) = "all content") Or (LCase(WorkingName) = "add-ons with no collection")
                If Not IsVisible Then
                    '
                    ' IsVisible if it is not in the EmptyNodeList (has child entries)
                    '
                    IsVisible = (InStr(1, "," & env.emptyNodeList & ",", "," & NavigatorID & ",") = 0)
                End If
                If IsVisible Then
                    '
                    ' hide the legacy node 'switch to navigator'
                    '
                    IsVisible = LCase(WorkingName) <> "switch to navigator"
                End If
                If IsVisible Then
                    Dim IconOpened As String = ""
                    Dim IconClosed As String = ""
                    Dim IconNoSubNodes As String = ""
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
                            Case "add-ons with no collection"
                                '
                                ' any Navigator Entry named 'all content' returns the content list
                                '
                                DivIDBase = CStr(NavigatorID)
                            Case "all content"
                                '
                                ' any Navigator Entry named 'all content' returns the content list
                                '
                                DivIDBase = CStr(NavigatorID)
                            Case Else
                                '
                                ' This entry is made from a navigator entry record
                                '
                                Select Case NodeType
                                    Case NodeTypeEnum.NodeTypeAddon
                                        '
                                        ' List of Addons
                                        '
                                        DivIDBase = "a" & NavigatorID
                                    Case NodeTypeEnum.NodeTypeCollection
                                        '
                                        ' List of Collections
                                        '
                                        DivIDBase = "c" & CollectionID
                                    Case NodeTypeEnum.NodeTypeContent
                                        '
                                        ' List of content
                                        '
                                        DivIDBase = "d" & NavigatorID
                                    Case Else
                                        '
                                        ' List of Navigator Entries
                                        '
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
                    Dim workingNameHtmlEncoded As String = ""
                    '
                    ' setup link
                    '
                    If Not String.IsNullOrEmpty(link) Then
                        NavLinkHTMLId = "n" & NavigatorID
                        workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                        If NewWindow Then
                            workingNameHtmlEncoded = "<a class=navDrag name=""navLink"" id=""" & NavLinkHTMLId & """ href=""" & link & """ target=""_blank"" title=""Open '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                        Else
                            workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & link & """ title=""Open '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                        End If
                    Else
                        hint = 1000
                        '
                        ' If link page, use this
                        '
                        If (addon IsNot Nothing) Then
                            hint = 1010
                            '
                            ' link to addon
                            '
                            link = ""
                            AddonGuid = addon.ccguid
                            AddonName = addon.name
                            If addon.remoteMethod Then
                                hint = 1020
                                NewWindow = True
                                '
                                ' -- url encode, but encode space as %20, not plus
                                For Each linkPart In AddonName.Split(" "c)
                                    link &= "%20" & cp.Utils.EncodeUrl(linkPart)
                                Next
                                link = link.Substring(3)
                                hint = 1030
                                '
                                If Not String.IsNullOrEmpty(link) Then
                                    link = link.Replace("\", "/")
                                    If link.Substring(0, 1) <> "/" Then
                                        link = "/" & link
                                    End If
                                End If
                                hint = 1040
                                'link = env.adminUrl & "?" & RequestNameRemoteMethodAddon & "=" & cp.Utils.EncodeRequestVariable(AddonName)
                            End If
                            If String.IsNullOrEmpty(link) Then
                                If Not String.IsNullOrEmpty(AddonGuid) Then
                                    link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "addonguid", AddonGuid, True)
                                Else
                                    link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "addonid", CStr(addon.id), True)
                                End If
                            End If
                            NavLinkHTMLId = "a" & addon.id
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                            If NewWindow Then
                                workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & link & """ target=""_blank"" title=""Run '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                            Else
                                workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & link & """ title=""Run '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                            End If
                            '
                            If (env.allowToolTip And (addon.template Or addon.content Or (addon.navTypeId > 1))) Then
                                '
                                ' -- add tooltip for addons
                                workingNameHtmlEncoded &= "<i data-tooltip=""" & AddonGuid & """ class=""contensiveToolTip fas fa-info-circle"" href=""#"" data-toggle=""tooltip"" data-html=""true"" rel=""tooltip"" title=""Click for details.""></i>"
                            End If
                        ElseIf ContentID <> 0 Then
                            hint = 2000
                            '
                            ' go edit the content
                            '
                            link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "cid", CStr(ContentID), True)
                            NavLinkHTMLId = "c" & ContentID
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                            If NewWindow Then
                                workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & link & """ target=""_blank"" title=""List All '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                            Else
                                workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & link & """ title=""List All '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                            End If
                        ElseIf HelpAddonID <> 0 Then
                            hint = 3000
                            '
                            ' go to Addon Help
                            '
                            link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "helpaddonid", CStr(HelpAddonID), True)
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                            If NewWindow Then
                                workingNameHtmlEncoded = "<a href=""" & link & """ target=""_blank"" title=""Help for Add-on '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                            Else
                                workingNameHtmlEncoded = "<a href=""" & link & """ title=""Help for Add-on '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                            End If
                        ElseIf helpCollectionID <> 0 Then
                            hint = 4000
                            '
                            ' go to Collection Help
                            '
                            Using cs13 As CPCSBaseClass = cp.CSNew
                                BlockNode = True
                                If cs13.Open("add-on collections", "id=" & helpCollectionID, "name", True, "name,helpLink,help") Then

                                    Dim collectionName As String = cs13.GetText("name")
                                    Dim collectionHelpLink As String = cs13.GetText("helpLink")
                                    Dim collectionHelp As String = cs13.GetText("help")
                                    '
                                    WorkingName = collectionName
                                    link = collectionHelpLink
                                    workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                                    If Not String.IsNullOrEmpty(link) Then
                                        BlockNode = False
                                        NewWindow = True
                                    ElseIf Not String.IsNullOrEmpty(collectionHelp) Then
                                        BlockNode = False
                                        link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "helpcollectionid", CStr(helpCollectionID), True)
                                    End If
                                    If Not BlockNode Then
                                        If NewWindow Then
                                            workingNameHtmlEncoded = "<a href=""" & link & """ target=""_blank"" title=""Help for Collection '" & workingNameHtmlEncoded & "'"">Help</a>"
                                        Else
                                            workingNameHtmlEncoded = "<a href=""" & link & """ title=""Help for Collection '" & workingNameHtmlEncoded & "'"">Help</a>"
                                        End If
                                    End If

                                End If
                                cs13.Close()
                            End Using
                        Else
                            hint = 5000
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                        End If
                    End If
                    hint = 6000
                    '
                    If Not BlockNode Then
                        hint = 6010
                        If BlockSubNodes Or (addon IsNot Nothing) Then
                            hint = 6020
                            '
                            ' This is a hardcoded item (like Add-on), it has no subnodes
                            '
                            result &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                        Else
                            hint = 6030
                            '
                            DivIDClosed = DivIDBase & "a"
                            DivIDOpened = DivIDBase & "b"
                            DivIDContent = DivIDBase & "c"
                            DivIDEmpty = DivIDBase & "d"

                            If (ContentID = 0) And (InStr(1, env.emptyNodeList & ",", "," & NodeIDString & ",") <> 0) Then
                                hint = 6040
                                '
                                ' In EmptyNodeList
                                '
                                result &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            ElseIf InStr(1, env.openNodeList & ",", "," & NodeIDString & ",") <> 0 Then
                                hint = 6050
                                Dim NodeNavigatorJS As String = ""
                                '
                                ' This node is open
                                '
                                SubNav = NodeListController.getNodeList(cp, env, NodeIDString, NodeNavigatorJS)
                                Return_NavigatorJS &= NodeNavigatorJS
                                If Not String.IsNullOrEmpty(SubNav) Then
                                    '
                                    ' display the subnav
                                    '
                                    result &= cr & "<div class=ccNavLink ID=" & DivIDClosed & " style=""display:none;""><A class=""ccNavClosed"" href=""#"" onclick=""AdminNavOpenClick('" & DivIDClosed & "','" & DivIDOpened & "','" & DivIDContent & "','" & NodeIDString & "','" & DivIDEmpty & "');return false;"">" & IconClosed & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                    result &= cr & "<div class=ccNavLink ID=" & DivIDOpened & "><A class=""ccNavOpened"" href=""#"" onclick=""AdminNavCloseClick('" & DivIDOpened & "','" & DivIDClosed & "','" & DivIDContent & "','" & NodeIDString & "');return false;"">" & IconOpened & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                    result = result _
                                        & cr & "<div class=ccNavLinkChild ID=" & DivIDContent & ">" _
                                        & kmaIndent(SubNav) _
                                        & cr & "</div>"
                                Else
                                    '
                                    ' it has a NO subnav
                                    '
                                    ' -- what happens if we don't display empty nodes
                                    'result &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                End If
                            Else
                                hint = 6060
                                '
                                ' This node is closed
                                '
                                result &= cr & "<div class=ccNavLink ID=" & DivIDClosed & " ><A class=""ccNavClosed"" href=""#"" onclick=""AdminNavOpenClick('" & DivIDClosed & "','" & DivIDOpened & "','" & DivIDContent & "','" & NodeIDString & "','','" & DivIDContent & "');return false;"">" & IconClosed & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                result &= cr & "<div class=ccNavLink ID=" & DivIDOpened & " style=""display:none;""><A class=""ccNavOpened"" href=""#"" onclick=""AdminNavCloseClick('" & DivIDOpened & "','" & DivIDClosed & "','" & DivIDContent & "','" & NodeIDString & "');return false;"">" & IconOpened & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                result &= cr & "<div class=""ccNavLink ccNavLinkEmpty"" ID=" & DivIDEmpty & " style=""display:none;"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                result &= cr & "<div class=ccNavLinkChild ID=" & DivIDContent & " style=""display:none;margin-left:20px;"">&nbsp;&nbsp;&nbsp;&nbsp;<img src=""https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/ajax-loader-small.gif"" width=""16"" height=""16""></div>"
                            End If
                        End If
                    End If
                End If
                '
                Return result
            Catch ex As Exception
                cp.Site.ErrorReport(ex, $"NodeController.getNode, hint [{hint}]")
                Throw
            End Try
        End Function
    End Class
End Namespace


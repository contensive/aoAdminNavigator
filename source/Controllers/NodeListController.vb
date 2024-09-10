
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Contensive.AdminNavigator
    Public NotInheritable Class NodeListController
        '        
        '====================================================================================================
        ''' <summary>
        ''' get Node list?
        ''' </summary>
        Friend Shared Function getNodeList(cp As CPBaseClass, env As ApplicationEnvironmentModel, ParentNode As String, ByRef Return_NavigatorJS As String) As String
            Dim returnNav As String = ""
            Dim hint As Integer = 0
            Try
                Dim TopParentNode As String = ParentNode
                Dim parentNodeStack() As String
                If String.IsNullOrEmpty(TopParentNode) Then
                    '
                    ' bad call
                    '
                    ReDim parentNodeStack(0)
                    parentNodeStack(0) = ""
                Else
                    '
                    ' load ParentNodes with argument
                    '
                    parentNodeStack = Split(TopParentNode, ".")
                End If
                Dim NodeNavigatorJS As String
                Dim ATag As String
                Dim FieldList As String
                Dim IconNoSubNodes As String
                Dim NavigatorID As Integer
                Dim CollectionID As Integer
                Dim Name As String
                Dim ContentID As Integer
                Dim Link As String
                Dim Criteria As String
                Dim BlockSubNodes As Boolean
                Dim NodeType As NodeTypeEnum
                Dim NavIconType As Integer
                Dim NavIconTitle As String = ""
                Dim NavIconTitleHtmlEncoded As String
                Dim ContentControlID As Integer
                Dim csChildList As CPCSBaseClass = cp.CSNew()

                Const AutoManageAddons = True
                Dim nodeIDString As String = ""
                Select Case parentNodeStack(0)
                        '
                        ' Open CS so:
                        '   Name = the caption that is displayed for the entry
                        '   ID (NavigatorID) = the NavigatorEntry the record represents
                        '       if the node has no navigation entry, NavigatorID=0 if there are no child nodes
                        '       this number is used for the open/close javascript, as well as stored to remember open/close state
                        '       during future hits, as the menu is built, this is checked in the open/close list for a match
                        '       NavigatorID=0 will only work if the node has not child nodes
                        '   AddonID = the ID of the addon that should be run if this entry is selected, 0 otherwise
                        '   CollectionID, if this is the manage add-ons section, this is the collection node
                        '   NewWindow = 0 or 1, if the link opens in new window
                        '   ContentID = the id of the content to be opened in list mode if the entry is clicked
                        '   Link = URL to link the menu entry
                        '   NodeIDString = unique string that represents this node
                        '       Navigator Entries - 'n'+EntryID
                        '       Collections = 'c'+CollectionID
                        '       Add-ons = 'a'+AddonID
                        '       CDefs = 'd'+ContentID
                        '
                    Case NodeIDManageAddons
                        '
                        ' Special Case: clicked on Manage Add-ons ("manageaddons")
                        ' Link to Add-on Manager
                        returnNav &= NodeManageAddonsController.getNode(cp, env, Return_NavigatorJS, nodeIDString)
                    Case NodeIDManageAddonsCollectionPrefix
                        '
                        ' Special Case: clicked on Manage Add-ons.collection
                        ' ParentNode(1) is the id of the collection they clicked on
                        ' List all add-ons
                        ' List all CDef
                        ' Add Collection Help
                        ' Add Layouts associated with collection
                        '
                        Dim nodeHtml As String = ""
                        Dim cacheName As String
                        CollectionID = 0
                        If UBound(parentNodeStack) > 0 Then
                            CollectionID = cp.Utils.EncodeInteger(parentNodeStack(1))
                        End If
                        cacheName = "addonNav." & NodeIDManageAddonsCollectionPrefix & "." & CollectionID & "." & cp.User.Id.ToString
                        nodeHtml = cp.Cache.GetText(cacheName)
                        If String.IsNullOrEmpty(nodeHtml) Then
                            '
                            ' Help Icon
                            '
                            Name = "Help"
                            nodeIDString = ""
                            NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle)
                            nodeHtml &= NodeController.getNode(cp, env, 0, 0, CollectionID, 0, 0, "", Nothing, 0, Name, env.emptyNodeList, 0, NavIconTypeHelp, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, True, env.openNodeList, nodeIDString, NodeNavigatorJS, "")
                            Return_NavigatorJS &= NodeNavigatorJS
                            '
                            ' List out add-ons in this collection
                            '
                            nodeIDString = ""
                            FieldList = "*"
                            '
                            ' -- block all addons not marked to be on the admin-nav (admin boolean)
                            ' -- big change, developers do not need this long list.
                            Criteria = "(collectionid=" & CollectionID & ")and(admin>0)"
                            Dim NameSuffix As String
                            For Each addon As AddonModel In DbBaseModel.createList(Of AddonModel)(cp, Criteria, "name")
                                Name = Trim(addon.name)
                                NameSuffix = ""
                                Dim linkSuffixList As String = ""
                                If env.isDeveloper Then
                                    linkSuffixList &= "<a href=""" & env.addonEditAddonUrlPrefix & addon.id & """>edit</a>"
                                    If Not addon.admin Then
                                        linkSuffixList &= ",dev"
                                    End If
                                    linkSuffixList = "&nbsp;(" & linkSuffixList & ")"
                                End If
                                Dim addonidx As Integer = addon.id
                                NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name)
                                ContentControlID = addon.contentControlId
                                Select Case addon.navTypeId
                                    Case 2
                                        NavIconType = NavIconTypeReport
                                    Case 3
                                        NavIconType = NavIconTypeSetting
                                    Case 4
                                        NavIconType = NavIconTypeTool
                                    Case Else
                                        NavIconType = NavIconTypeAddon
                                End Select
                                nodeHtml &= NodeController.getNode(cp, env, 0, ContentControlID, 0, 0, 0, "", addon, 0, Name, env.emptyNodeList, 0, NavIconType, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeAddon, False, False, env.openNodeList, nodeIDString, NodeNavigatorJS, linkSuffixList)
                                Return_NavigatorJS &= NodeNavigatorJS

                            Next
                            '
                            ' List out cdefs connected to this collection
                            '
                            nodeIDString = ""
                            Criteria = "(collectionid=" & CollectionID & ")"
                            If env.isDeveloper Then
                            ElseIf cp.User.IsAdmin Then
                                Criteria &= "and(developeronly=0)"
                            Else
                                Criteria &= "and(developeronly=0)and(adminonly=0)"
                            End If
                            Dim LastContentID As Integer
                            Dim DupsFound As Boolean
                            LastContentID = -1
                            DupsFound = False
                            Dim SQL As String = "select c.id,c.name,c.contentcontrolid,c.developeronly,c.adminonly from ccContent c left join ccAddonCollectionCDefRules r on r.contentid=c.id where " & Criteria & " order by c.name"
                            Dim cs7 As CPCSBaseClass = cp.CSNew()
                            If cs7.OpenSQL(SQL) Then
                                Do
                                    Name = Trim(cs7.GetText("name"))
                                    ContentID = cs7.GetInteger("id")
                                    If ContentID = LastContentID Then
                                        DupsFound = True
                                    Else
                                        Dim linkSuffixList As String = ""
                                        If env.isDeveloper Then
                                            linkSuffixList = "<a href=""" & env.contentFieldEditToolPrefix & ContentID & """>edit</a>"
                                            If cs7.GetBoolean("developeronly") Then
                                                linkSuffixList &= ",dev"
                                            ElseIf cs7.GetBoolean("adminonly") Then
                                                linkSuffixList &= ",adm"
                                            End If
                                            If Not String.IsNullOrEmpty(linkSuffixList) Then
                                                linkSuffixList = "&nbsp;(" & linkSuffixList & ")"
                                            End If
                                        End If
                                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name)
                                        ContentControlID = cs7.GetInteger("ContentControlID")
                                        NameSuffix = ""
                                        nodeHtml &= NodeController.getNode(cp, env, 0, ContentControlID, 0, 0, ContentID, "", Nothing, 0, Name, env.emptyNodeList, 0, NavIconTypeContent, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeContent, False, True, env.openNodeList, nodeIDString, NodeNavigatorJS, linkSuffixList)
                                        Return_NavigatorJS &= NodeNavigatorJS
                                    End If
                                    LastContentID = ContentID
                                    Call cs7.GoNext()
                                Loop While cs7.OK()
                            End If
                            Call cs7.Close()
                            '
                            ' list all data records associated to this collection
                            '
                            Dim dataRecordList As String = ""
                            Dim dataRecords As New List(Of String)
                            Dim dataRecordParts As String()
                            'Dim dataRecordId As Integer
                            'Dim dataRecordGuid As String
                            Dim dataRecordName As String
                            Dim dataRecordCdefName As String
                            Dim dataRecordCdefID As Integer

                            If cs7.Open("add-on collections", "id=" & CollectionID) Then
                                dataRecordList = cs7.GetText("dataRecordList")
                            End If
                            Call cs7.Close()
                            hint = 1100
                            If Not String.IsNullOrEmpty(dataRecordList) Then
                                Try
                                    dataRecords.AddRange(dataRecordList.Split(vbCrLf.ToCharArray))
                                    For Each dataRecord As String In dataRecords
                                        dataRecordParts = dataRecord.Split(",".ToCharArray)
                                        dataRecordCdefName = dataRecordParts(0)
                                        dataRecordCdefID = cp.Content.GetID(dataRecordCdefName)
                                        If dataRecordCdefID = 0 Then
                                            ' -- invalid content, skip it
                                        Else
                                            ' -- export records from this content
                                            Dim sqlCriteria As String = ""
                                            If String.IsNullOrEmpty(dataRecordCdefName) Then
                                                ' -- no comma, export all records
                                            Else
                                                If dataRecordParts.Length >= 2 Then
                                                    hint = 1110
                                                    ' 
                                                    ' contentname,(name or guid)
                                                    If String.IsNullOrEmpty(dataRecordParts(1)) Then
                                                        ' -- no data given, list all records in export
                                                    Else
                                                        If dataRecordParts(1).Substring(0, 1) = "{" Then
                                                            sqlCriteria = "ccguid=" & cp.Db.EncodeSQLText(dataRecordParts(1))
                                                        Else
                                                            sqlCriteria = "name=" & cp.Db.EncodeSQLText(dataRecordParts(1))
                                                        End If
                                                    End If
                                                    hint = 1120
                                                End If
                                            End If
                                            If cs7.Open(dataRecordCdefName, sqlCriteria) Then
                                                Do
                                                    dataRecordName = cs7.GetText("name")


                                                    NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Edit '" & dataRecordName & "' in '" & dataRecordCdefName & "'")
                                                    IconNoSubNodes = IconRecord
                                                    IconNoSubNodes = Replace(IconNoSubNodes, "{title}", NavIconTitleHtmlEncoded)
                                                    Link = "?id=" & cs7.GetInteger("id").ToString & "&cid=" & dataRecordCdefID.ToString & "&af=4"
                                                    ATag = "<a href=""" & Link & """ title=""" & NavIconTitleHtmlEncoded & """>"
                                                    nodeHtml &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & ATag & IconNoSubNodes & "</a>&nbsp;" & ATag & dataRecordCdefName & ":" & dataRecordName & "</a></div>"
                                                    'nodeHtml &= GetNode(cp, env, 0, 0, 0, 0, 0, "/admin?af=4&cid=" & dataRecordCdefID.ToString & "&id=" & cs7.GetInteger("id"), 0, 0, dataRecordCdefName & ":" & cs7.GetText("name"), env.EmptyNodeList, 0, NavIconTypeContent, "Data", AutoManageAddons, NodeTypeEnum.NodeTypeContent, False, True, env.OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                                                    Call cs7.GoNext()
                                                Loop While cs7.OK

                                            End If
                                            Call cs7.Close()
                                        End If
                                    Next
                                    hint = 1100
                                    '
                                    If DupsFound Then
                                        SQL = "select b.id from ccAddonCollectionCDefRules a,ccAddonCollectionCDefRules b where (a.id<b.id) and (a.contentid=b.contentid) and (a.collectionid=b.collectionid)"
                                        SQL = "delete from ccAddonCollectionCDefRules where id in (" & SQL & ")"
                                        Call cp.Db.ExecuteNonQuery(SQL)
                                    End If
                                    Call cp.Cache.Store(cacheName, nodeHtml, Now.AddHours(1), env.cacheDependencyList)
                                Catch ex As Exception
                                    cp.Site.ErrorReport(ex, $"Error creating nodes for Data Records, continuing")
                                End Try
                            End If
                        End If
                        returnNav &= nodeHtml
                    Case NodeIDManageAddonsAdvanced
                        '
                        ' Special Case: clicked on Manage Add-ons.advanced
                        '   edit links for Add-ons, Add-on Collections
                        '
                        ' Folder to Add-ons without Collections
                        '
                        nodeIDString = NodeIDAddonsNoCollection
                        returnNav &= NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "", Nothing, 0, "Add-ons With No Collection", env.emptyNodeList, 0, NavIconTypeAddon, "Add-ons With No Collection", AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, env.openNodeList, nodeIDString, NodeNavigatorJS, "")
                        Return_NavigatorJS &= NodeNavigatorJS
                        '
                        Name = "Add-ons"
                        returnNav &= NodeController.getNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", Nothing, 0, Name, env.emptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, env.openNodeList, "", NodeNavigatorJS, "")
                        Return_NavigatorJS &= NodeNavigatorJS
                        '
                        Name = "Add-on Collections"
                        returnNav &= NodeController.getNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", Nothing, 0, Name, env.emptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, env.openNodeList, "", NodeNavigatorJS, "")
                        Return_NavigatorJS &= NodeNavigatorJS
                        '
                        Name = "Add-on Categories"
                        returnNav &= NodeController.getNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", Nothing, 0, Name, env.emptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, env.openNodeList, "", NodeNavigatorJS, "")
                        Return_NavigatorJS &= NodeNavigatorJS
                        '
                    Case NodeIDAddonsNoCollection
                        '
                        ' special case: Add-on List that do not have collections
                        '
                        CollectionID = 0

                        FieldList = "0 as ContentControlID,A.Name as Name,A.ID as ID,A.ID as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeAddon & " as NavIconType,A.Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
                        Dim SQL As String = "select" _
                            & " " & FieldList _
                            & " from ccAggregateFunctions A" _
                            & " left join ccAddonCollections C on C.ID=A.CollectionID" _
                            & " where C.ID is null" _
                            & " order by A.Name"
                        NodeType = NodeTypeEnum.NodeTypeAddon
                        BlockSubNodes = True
                        nodeIDString = ""
                        For Each addon As AddonModel In DbBaseModel.createList(Of AddonModel)(cp, "id in (select a.id from ccAggregateFunctions a left join ccAddonCollections C on C.ID=A.CollectionID where c.id is null)")
                            returnNav &= NodeController.getNode(cp, env, 0, addon.contentControlId, 0, 0, 0, "", addon, 0, Trim(addon.name), env.emptyNodeList, 0, NavIconTypeAddon, cp.Utils.EncodeHTML(addon.name), AutoManageAddons, NodeTypeEnum.NodeTypeAddon, False, False, env.openNodeList, nodeIDString, NodeNavigatorJS, "")
                            Return_NavigatorJS &= NodeNavigatorJS
                        Next
                    Case NodeIDLegacyMenu
                        '
                        ' Special Case: build old top menus under this Navigator entry
                        '
                        BlockSubNodes = False
                        Dim SQL As String = MenuSqlController.getMenuSQL(cp, "(parentid=0)or(parentid is null)", LegacyMenuContentName)
                        If Not csChildList.OpenSQL(SQL) Then
                            '
                            ' Empty list, add to env.EmptyNodeList
                            '
                            env.emptyNodeList &= "," & TopParentNode
                        End If
                    Case NodeIDAllContentList
                        '
                        ' special case: all content
                        '
                        FieldList = "Name,ID,0 as AddonID,0 as NewWindow,ID as ContentID,'' as LinkPage," & NavIconTypeContent & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid,0 as collectionid"
                        Dim SQL As String = "select " & FieldList & " from cccontent order by name"
                        csChildList.OpenSQL(SQL)
                        NodeType = NodeTypeEnum.NodeTypeContent
                        BlockSubNodes = True
                    Case "0", ""
                        '
                        ' Navigator Entries, list home(s) plus all roots
                        '
                        NodeType = NodeTypeEnum.NodeTypeEntry
                        BlockSubNodes = False
                        Link = cp.Utils.EncodeHTML("http://" & cp.Site.Domain)
                        returnNav &= cr & "<div class=ccNavLink><A href=""" & Link & """>" & IconPublicHome & "</A>&nbsp;<A href=""" & Link & """>Public Home</A></div>"
                        Link = cp.Utils.EncodeHTML(env.adminUrl)
                        returnNav &= cr & "<div class=ccNavLink><A href=""" & Link & """>" & IconAdminHome & "</A>&nbsp;<A href=""" & Link & """>Admin Home</A></div>"
                        Dim cs8 As CPCSBaseClass = cp.CSNew()
                        If cs8.OpenSQL(MenuSqlController.getMenuSQL(cp, "((Parentid=0)or(Parentid is null))", NavigatorContentName)) Then
                            Do
                                Name = Trim(cs8.GetText("name"))
                                NavigatorID = cs8.GetInteger("ID")
                                NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name)
                                nodeIDString = CStr(NavigatorID)
                                If AutoManageAddons Then
                                    '
                                    ' special cases - root nodes that do not just deliver menu entries
                                    '
                                    Select Case LCase(Name)
                                        Case "manage add-ons"
                                            nodeIDString = NodeIDManageAddons
                                        Case "settings"
                                            nodeIDString = NodeIDSettings
                                        Case "tools"
                                            nodeIDString = NodeIDTools
                                        Case "reports"
                                            nodeIDString = NodeIDReports
                                    End Select
                                End If
                                returnNav &= NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "", Nothing, 0, Name, env.emptyNodeList, NavigatorID, NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, env.openNodeList, nodeIDString, NodeNavigatorJS, "")
                                Return_NavigatorJS &= NodeNavigatorJS
                                Call cs8.GoNext()
                            Loop While cs8.OK
                        End If
                        Call cs8.Close()
                        '
                        ' Add a Legacy Menu node to the top-most parent menu at the very end
                        '
                        If cp.Utils.EncodeBoolean(cp.Site.GetText("AllowNavigatorLegacyEntry", "0")) Then
                            Name = "Legacy Menu"
                            NavIconTitleHtmlEncoded = "Legacy Menu"
                            nodeIDString = NodeIDLegacyMenu
                            returnNav &= NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "", Nothing, 0, Name, env.emptyNodeList, 0, NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, BlockSubNodes, env.openNodeList, nodeIDString, NodeNavigatorJS, "")
                            Return_NavigatorJS &= NodeNavigatorJS
                        End If
                    Case NodeIDSettings
                        '
                        ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                        '
                        returnNav &= NodeListMixedController.getNodeListMixed(cp, env, env.emptyNodeList, TopParentNode, AutoManageAddons, NodeType, env.openNodeList, 3, cp.Content.GetRecordID("navigator entries", "settings"), NavIconTypeSetting, NodeNavigatorJS)
                        Return_NavigatorJS &= NodeNavigatorJS
                    Case NodeIDTools
                        '
                        ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                        '
                        returnNav &= NodeListMixedController.getNodeListMixed(cp, env, env.emptyNodeList, TopParentNode, AutoManageAddons, NodeType, env.openNodeList, 4, cp.Content.GetRecordID("navigator entries", "tools"), NavIconTypeTool, NodeNavigatorJS)
                        Return_NavigatorJS &= NodeNavigatorJS
                    Case NodeIDReports
                        '
                        ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                        '
                        returnNav &= NodeListMixedController.getNodeListMixed(cp, env, env.emptyNodeList, TopParentNode, AutoManageAddons, NodeType, env.openNodeList, 2, cp.Content.GetRecordID("navigator entries", "reports"), NavIconTypeReport, NodeNavigatorJS)
                        Return_NavigatorJS &= NodeNavigatorJS
                    Case Else
                        '
                        ' numeric node (default case) - list navigator records with parent=TopParentNode
                        '
                        Dim CS As Integer = -1
                        If IsNumeric(TopParentNode) Then
                            If InStr(1, env.emptyNodeList & ",", "," & TopParentNode & ",") <> 0 Then
                                '
                            Else
                                '
                                ' Navigator Entries, child under TopParentNode
                                '
                                Dim SQL As String = MenuSqlController.getMenuSQL(cp, "parentid=" & TopParentNode, "")
                                BlockSubNodes = False
                                If Not csChildList.OpenSQL(SQL) Then
                                    '
                                    ' Empty list, add to env.EmptyNodeList
                                    '
                                    env.emptyNodeList &= "," & TopParentNode
                                End If
                            End If
                        End If
                End Select
                '
                ' ----- List Navigator Nodes, if not already displayed
                '
                If (Not csChildList.OK) And (NodeType = NodeTypeEnum.NodeTypeEntry) Then
                    '
                    ' No child nodes, if this node includes a CID, list the first 20 content records with a 'more'
                    '
                    ContentID = 0
                    If IsNumeric(TopParentNode) Then
                        Call csChildList.Close()
                        If csChildList.Open(NavigatorContentName, "id=" & TopParentNode) Then
                            ContentID = csChildList.GetInteger("ContentID")
                        End If
                        If ContentID <> 0 Then
                            ContentID = ContentID
                            Dim ContentName As String = cp.Content.GetRecordName("content", ContentID)
                            If Not String.IsNullOrEmpty(ContentName) Then
                                csChildList.Close()
                                Dim Ptr As Integer = 0
                                If csChildList.Open(ContentName, "", "name", True, "ID,Name,ContentControlID", 20, 1) Then
                                    env.emptyNodeList = Replace(env.emptyNodeList, "," & TopParentNode, "")
                                    Do
                                        NavigatorID = csChildList.GetInteger("ID")
                                        Dim RecordName As String = csChildList.GetText("Name")
                                        If String.IsNullOrEmpty(RecordName) Then
                                            RecordName = "Record " & NavigatorID
                                        End If
                                        '
                                        If Len(RecordName) > 53 Then
                                            RecordName = Left(RecordName, 25) & "..." & Right(RecordName, 25)
                                        End If
                                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Edit '" & RecordName & "' in '" & ContentName & "'")
                                        IconNoSubNodes = IconRecord
                                        IconNoSubNodes = Replace(IconNoSubNodes, "{title}", NavIconTitleHtmlEncoded)
                                        Link = "?id=" & NavigatorID & "&cid=" & csChildList.GetInteger("ContentControlID") & "&af=4"
                                        ATag = "<a href=""" & Link & """ title=""" & NavIconTitleHtmlEncoded & """>"
                                        returnNav &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & ATag & IconNoSubNodes & "</a>&nbsp;" & ATag & RecordName & "</a></div>"
                                        Ptr += 1
                                        Call csChildList.GoNext()
                                    Loop While csChildList.OK() And Ptr < 20
                                    If Ptr = 20 Then
                                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Open All '" & NavigatorContentName & "'")
                                        Link = "?cid=" & ContentID
                                        Dim IconClosed As String
                                        returnNav &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconClosed & "&nbsp;<a href=""" & Link & """ title=""" & NavIconTitleHtmlEncoded & """>more...</a></div>"
                                    End If
                                End If
                            End If
                        End If
                    End If
                ElseIf csChildList.OK Then
                    '
                    ' List out child menus
                    '
                    Do
                        Dim addon As AddonModel = DbBaseModel.create(Of AddonModel)(cp, csChildList.GetInteger("AddonID"))
                        '
                        CollectionID = csChildList.GetInteger("CollectionID")
                        NavigatorID = csChildList.GetInteger("ID")
                        Name = Trim(csChildList.GetText("name"))
                        Dim NewWindow As Boolean = csChildList.GetBoolean("newwindow")
                        ContentID = csChildList.GetInteger("ContentID")
                        Link = Trim(csChildList.GetText("LinkPage"))
                        NavIconType = csChildList.GetInteger("NavIconType")
                        NavIconTitle = csChildList.GetText("NavIconTitle")
                        Dim HelpAddonID As Integer = csChildList.GetInteger("HelpAddonID")
                        If HelpAddonID <> 0 Then
                            HelpAddonID = HelpAddonID
                        End If
                        Dim helpCollectionID As Integer = csChildList.GetInteger("HelpCollectionID")
                        If String.IsNullOrEmpty(NavIconTitle) Then
                            NavIconTitle = Name
                        End If
                        ContentControlID = csChildList.GetInteger("ContentControlID")
                        If LCase(Name) = "all content" Then
                            '
                            ' special case: any Navigator Entry named 'all content' returns the content list
                            '
                            nodeIDString = NodeIDAllContentList
                        Else
                            nodeIDString = CStr(NavigatorID)
                        End If
                        Dim linkSuffixList As String = ""
                        If (ContentID <> 0) And (env.isDeveloper) Then
                            linkSuffixList = "<a href=""" & env.contentFieldEditToolPrefix & ContentID & """>edit</a>"
                            If Not String.IsNullOrEmpty(linkSuffixList) Then
                                linkSuffixList = "&nbsp;(" & linkSuffixList & ")"
                            End If
                        End If
                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle)
                        Dim SettingPageID As Integer
                        returnNav &= NodeController.getNode(cp, env, CollectionID, ContentControlID, helpCollectionID, HelpAddonID, ContentID, Link, addon, SettingPageID, Name, env.emptyNodeList, NavigatorID, NavIconType, NavIconTitleHtmlEncoded, AutoManageAddons, NodeType, NewWindow, BlockSubNodes, env.openNodeList, nodeIDString, NodeNavigatorJS, linkSuffixList)
                        Return_NavigatorJS &= NodeNavigatorJS
                        Call csChildList.GoNext()
                    Loop While csChildList.OK
                    Call csChildList.Close()
                End If
                Return returnNav
            Catch ex As Exception
                cp.Site.ErrorReport(ex, $"hint [{hint}]")
                Throw
            End Try
        End Function
    End Class
End Namespace


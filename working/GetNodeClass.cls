VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetNodeClass"
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
    Dim ParentNode As String
    Dim OpenNodeList As String
    Dim NavigatorJS As String
    '
    Set Main = MainObject
    Set Csv = CsvObject
    '
    '
    ' For Page Addons, return the result
    '
    
    OpenNodeList = Main.GetVisitProperty("AdminNavOpenNodeList", "")
    ParentNode = Main.GetStreamText("nodeid")
    If ParentNode <> "" Then
        OpenNodeList = OpenNodeList & "," & ParentNode
        Call Main.SetVisitProperty("AdminNavOpenNodeList", OpenNodeList)
    End If
    Execute = GetNavigator(ParentNode, OpenNodeList, NavigatorJS)
    If NavigatorJS <> "" Then
        NavigatorJS = "" _
            & "if(window.navDrop) {" _
            & NavigatorJS _
            & "};"
        Call Main.AddHeadScriptCode(NavigatorJS, "Admin Navigator")
        'Call Main.AddHeadScriptCode("alert('addheadscriptcode');", "Admin Navigator")
    End If
    '
    Exit Function
ErrorTrap:
    'HandleError
End Function
'
'
'
Friend Function GetNavigator(ParentNode As String, OpenNodeList As String, Return_NavigatorJS) As String
    On Error GoTo ErrorTrap
    '
    Const AutoManageAddons = True
    '
    Dim NodeNavigatorJS As String
    Dim ATag As String
    Dim Index As New FastIndex4.FastIndexClass
    Dim SortNodes() As SortNodeType
    Dim SortPtr As Long
    Dim SortCnt As Long
    Dim test As String
    Dim NameSuffix As String
    Dim NodeID As Long
    Dim CSAddon As Long
    Dim SettingPageID As Long
    Dim FieldList As String
    Dim IsVisible As Boolean
    Dim BakeName As String
    '
    ' ----- Timer Trace
    Dim TimerTraceMarker As String
    Dim StartTickCount As Long
    'StartTickCount = GetTickCount
    'TimerTraceMarker = Main.GetRandomLong
    'Main.ToolsPanelTimerTrace = Main.ToolsPanelTimerTrace & CR & "<li class=""ccAdminSmall ccPanel"">AdminClass.GetNavigator(): " & TimerTraceMarker & " msec</li>" & CR & "<ul>"
    ' ----- /Timer Trace
    '
    Dim NodeIDString As String
    Dim IconNoSubNodes As String
    Dim IconOpened As String
    Dim IconClosed As String
    Dim NavigatorID As Long
    'Dim DivIDBase As String
    Dim SQL As String
    Dim CollectionID As Long
    Dim DivIDClosed As String
    Dim DivIDOpened As String
    Dim DivIDContent As String
    Dim DivIDEmpty As String
    Dim NodeSN As String
    Dim CS As Long
    Dim s As String
    Dim EntryID As Long
    Dim Name As String
    Dim Caption As String
    Dim SubNav As String
    Dim NewWindow As Boolean
    Dim ContentID As Long
    Dim HelpAddonID As Long
    Dim helpCollectionID As Long
    Dim addonid As Long
    Dim Link As String
    Dim Criteria As String
    Dim BlockSubNodes As Boolean
    Dim TopParentNode As String
    Dim ParentNodes() As String
    'Dim NodeIDStringPrefix As String
    'Dim NodeIDString As String
    Dim NodeType As NodeTypeEnum
    Dim ContentName As String
    Dim ContentTableName As String
    Dim Ptr As Long
    Dim RecordName As String
    Dim NavIconType As String
    Dim NavIconTitle As String
    Dim NavIconTitleHtmlEncoded As String
    Dim EmptyNodeList As String
    Dim EmptyNodeListInitial As String
    Dim RS As Recordset
    Dim EmptyNodes() As Long
    Dim Temp As Variant
    Dim LegacyMenuControlID As Long
    Dim IsLegacyMenuRootNode As Boolean
    Dim ContentControlID As Long
    Dim BuildVersion As String
    Dim isDeveloper As Boolean
    '
    BuildVersion = Main.SiteProperty_BuildVersion
    isDeveloper = Main.isDeveloper()
    If BuildVersion < "3.4.175" Then
        GetNavigator = "Upgrade your site Database to support this feature."
    Else
        If isDeveloper Then
            BakeName = "AdminNav EmptyNodeList Dev"
        ElseIf Main.IsAdmin() Then
            BakeName = "AdminNav EmptyNodeList Admin"
        Else
            BakeName = "AdminNav EmptyNodeList CM" & Main.memberID
        End If
        EmptyNodeList = Main.ReadBake(BakeName)
        If EmptyNodeList = "" Then
            SQL = "select n.ID from ccMenuEntries n left join ccMenuEntries c on c.parentid=n.id Where c.ID Is Null group by n.id"
            'SQL = "select n.ID from ccMenuEntries n left join ccMenuEntries c on c.parentid=n.id Where c.ID Is Null And n.ContentControlID=" & Main.GetContentID("Navigator Entries") & " group by n.id"
            Set RS = Main.executesql("default", SQL)
            If Not (RS Is Nothing) Then
                If Not RS.EOF Then
                    EmptyNodeList = RS.GetString(adClipString, , "", ",")
                    'EmptyNodeList = Join(temp(0), ",")
                End If
            End If
            Set RS = Nothing
            Call Main.SaveBake(BakeName, EmptyNodeList, "Navigator Entries")
        End If
        EmptyNodeListInitial = EmptyNodeList
        TopParentNode = ParentNode
        If TopParentNode = "" Then
            '
            ' bad call
            '
            ReDim ParentNodes(0)
            ParentNodes(0) = ""
        Else
            '
            ' load ParentNodes with argument
            '
            ParentNodes = Split(TopParentNode, ".")
        End If
        LegacyMenuControlID = Main.GetContentID("Menu Entries")
        Select Case ParentNodes(0)
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
                '
                NodeIDString = ""
                addonid = Main.GetRecordID("Add-ons", "Add-on Manager")
                s = s & GetNavigatorNode(0, 0, 0, 0, 0, "?addonguid=" & AddonManagerGuid, addonid, 0, "Add-on Manager", LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeAddon, "Add-on Manager", AutoManageAddons, NodeTypeAddon, 0, False, OpenNodeList, NodeIDString, NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                '
                ' List Collections
                '
                FieldList = "Name,0 as id,ccaddoncollections.id as collectionid,0 as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeFolder & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid"
                'FieldList = "Name,id as collectionid,0 as ID,0 as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeFolder & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid"
                Criteria = "((system=0)or(system is null))"
                If (BuildVersion >= "4.1.512") And (Not isDeveloper) Then
                    Criteria = Criteria & "and((blockNavigatorNode=0)or(blockNavigatorNode is null))"
                End If
                CS = Main.openCSContent("Add-on Collections", Criteria, , , , , FieldList)
                NodeType = NodeTypeCollection
                BlockSubNodes = False
                Do While Main.iscsok(CS)
                    Name = Trim(Main.getcsText(CS, "name"))
                    NavIconTitle = Name
                    CollectionID = Main.GetCSInteger(CS, "collectionid")
                    NodeIDString = NodeIDManageAddonsCollectionPrefix & "." & CollectionID
                    NavIconTitleHtmlEncoded = kmaEncodeHTML(NavIconTitle)
                    s = s & GetNavigatorNode(CollectionID, 0, 0, 0, 0, "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeAddon, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeCollection, 0, False, OpenNodeList, NodeIDString, NodeNavigatorJS)
                    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    Call Main.NextCSRecord(CS)
                Loop
                Call Main.closeCs(CS)
                '
                ' Advanced folder to contain edit links to create addons and collections
                '
                NodeIDString = NodeIDManageAddonsAdvanced
                s = s & GetNavigatorNode(0, 0, 0, 0, 0, "", 0, 0, "Advanced", LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeFolder, "Add-ons With No Collection", AutoManageAddons, NodeTypeEntry, 0, False, OpenNodeList, NodeIDString, NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
            Case NodeIDManageAddonsCollectionPrefix
                '
                ' Special Case: clicked on Manage Add-ons.collection
                ' ParentNode(1) is the id of the collection they clicked on
                ' List all add-ons
                ' List all CDef
                ' Add Collection Help
                '
                CollectionID = 0
                If UBound(ParentNodes) > 0 Then
                    CollectionID = kmaEncodeInteger(ParentNodes(1))
                End If
                '
                ' Help Icon
                '
                Name = "Help"
                NodeIDString = ""
                NavIconTitleHtmlEncoded = kmaEncodeHTML(NavIconTitle)
                s = s & GetNavigatorNode(0, 0, CollectionID, 0, 0, "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeHelp, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEntry, 0, False, OpenNodeList, NodeIDString, NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                '
                ' List out add-ons in this collection
                '
                NodeIDString = ""
                FieldList = "*"
                'FieldList = "Name,id as collectionid,0 as ID,0 as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeFolder & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid"
                Criteria = "(collectionid=" & CollectionID & ")"
                If Not isDeveloper Then
                    Criteria = Criteria & "and(admin<>0)"
                    'Criteria = Criteria & "and((template<>0)or(page<>0)or(admin<>0))"
                End If
                CS = Main.openCSContent("Add-ons", Criteria, , , , , FieldList)
                Do While Main.iscsok(CS)
                    Name = Trim(Main.getcsText(CS, "name"))
                    NameSuffix = ""
                    If isDeveloper Then
                        If Main.GetCSBoolean(CS, "content") Then
                            NameSuffix = NameSuffix & "c"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        If Main.GetCSBoolean(CS, "template") Then
                            NameSuffix = NameSuffix & "t"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        'Name = Name & "(" & NameSuffix & ")"
                        If Main.GetCSBoolean(CS, "email") Then
                            NameSuffix = NameSuffix & "m"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        If Main.GetCSBoolean(CS, "admin") Then
                            NameSuffix = NameSuffix & "n"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        If Main.GetCSBoolean(CS, "remotemethod") Then
                            NameSuffix = NameSuffix & "r"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        If Main.GetCSBoolean(CS, "onpagestartevent") Then
                            NameSuffix = NameSuffix & "b"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        If Main.GetCSBoolean(CS, "onpageendevent") Then
                            NameSuffix = NameSuffix & "a"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        If Main.GetCSBoolean(CS, "onbodystart") Then
                            NameSuffix = NameSuffix & "s"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        If Main.GetCSBoolean(CS, "onpageendevent") Then
                            NameSuffix = NameSuffix & "e"
                        Else
                            NameSuffix = NameSuffix & "-"
                        End If
                        Name = Name & " (" & NameSuffix & ")"
                    End If
                    addonid = Main.GetCSInteger(CS, "ID")
                    NavIconTitleHtmlEncoded = kmaEncodeHTML(Name)
                    ContentControlID = Main.GetCSInteger(CS, "ContentControlID")
                    Select Case Main.GetCSInteger(CS, "navtypeid")
                        Case 2
                            NavIconType = NavIconTypeReport
                        Case 3
                            NavIconType = NavIconTypeSetting
                        Case 4
                            NavIconType = NavIconTypeTool
                        Case Else
                            NavIconType = NavIconTypeAddon
                    End Select
                    s = s & GetNavigatorNode(0, ContentControlID, 0, 0, 0, "", addonid, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconType, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeAddon, 0, False, OpenNodeList, NodeIDString, NodeNavigatorJS)
                    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    Call Main.NextCSRecord(CS)
                Loop
                Call Main.closeCs(CS)
                '
                ' List out cdefs connected to this collection
                '
                NodeIDString = ""
                Criteria = "(collectionid=" & CollectionID & ")"
                If isDeveloper Then
                ElseIf Main.IsAdmin Then
                    Criteria = Criteria & "and(developeronly=0)"
                Else
                    Criteria = Criteria & "and(developeronly=0)and(adminonly=0)"
                End If
Dim LastContentID As Long
Dim DupsFound As Boolean
                LastContentID = -1
                DupsFound = False
                SQL = "select c.id,c.name,c.contentcontrolid,c.developeronly,c.adminonly from ccContent c left join ccAddonCollectionCDefRules r on r.contentid=c.id where " & Criteria & " order by c.name"
                CS = Main.OpenCSSQL("default", SQL)
                Do While Main.iscsok(CS)
                    Name = Trim(Main.getcsText(CS, "name"))
                    ContentID = Main.GetCSInteger(CS, "id")
                    If ContentID = LastContentID Then
                        DupsFound = True
                    Else
                        NavIconTitleHtmlEncoded = kmaEncodeHTML(Name)
                        ContentControlID = Main.GetCSInteger(CS, "ContentControlID")
                        NameSuffix = ""
                        If isDeveloper Then
                            If Main.GetCSBoolean(CS, "developeronly") Then
                                NameSuffix = NameSuffix & "--"
                            Else
                                If Main.GetCSBoolean(CS, "adminonly") Then
                                    NameSuffix = NameSuffix & "-a"
                                Else
                                    NameSuffix = NameSuffix & "ca"
                                End If
                            End If
                            
                            Name = Name & " (" & NameSuffix & ")"
                        End If
                        s = s & GetNavigatorNode(0, ContentControlID, 0, 0, ContentID, "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeContent, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeContent, 0, True, OpenNodeList, NodeIDString, NodeNavigatorJS)
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    End If
                    LastContentID = ContentID
                    Call Main.NextCSRecord(CS)
                Loop
                Call Main.closeCs(CS)
                If DupsFound Then
                    SQL = "select b.id from ccAddonCollectionCDefRules a,ccAddonCollectionCDefRules b where (a.id<b.id) and (a.contentid=b.contentid) and (a.collectionid=b.collectionid)"
                    SQL = "delete from ccAddonCollectionCDefRules where id in (" & SQL & ")"
                    Call Main.executesql("default", SQL)
                End If
            Case NodeIDManageAddonsAdvanced
                '
                ' Special Case: clicked on Manage Add-ons.advanced
                '   edit links for Add-ons, Add-on Collections
                '
                ' Folder to Add-ons without Collections
                '
                NodeIDString = NodeIDAddonsNoCollection
                s = s & GetNavigatorNode(0, 0, 0, 0, 0, "", 0, 0, "Add-ons With No Collection", LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeAddon, "Add-ons With No Collection", AutoManageAddons, NodeTypeEntry, 0, False, OpenNodeList, NodeIDString, NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                '
                Name = "Add-ons"
                s = s & GetNavigatorNode(0, 0, 0, 0, Main.GetContentID(Name), "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEntry, 0, False, OpenNodeList, "", NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                '
                Name = "Add-on Collections"
                s = s & GetNavigatorNode(0, 0, 0, 0, Main.GetContentID(Name), "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEntry, 0, False, OpenNodeList, "", NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                '
                Name = "Aggregate Access"
                s = s & GetNavigatorNode(0, 0, 0, 0, Main.GetContentID(Name), "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEntry, 0, False, OpenNodeList, "", NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                '
                Name = "Scripting Modules"
                s = s & GetNavigatorNode(0, 0, 0, 0, Main.GetContentID(Name), "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEntry, 0, False, OpenNodeList, "", NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
            Case NodeIDAddonsNoCollection
                '
                ' special case: Add-on List that do not have collections
                '
                CollectionID = 0
                If BuildVersion > "3.4.044" Then
                    FieldList = "0 as ContentControlID,A.Name as Name,A.ID as ID,A.ID as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeAddon & " as NavIconType,A.Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
                ElseIf BuildVersion > "3.3.982" Then
                    FieldList = "0 as ContentControlID,A.Name as Name,A.ID as ID,A.ID as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeAddon & " as NavIconType,A.Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
                Else
                    FieldList = "0 as ContentControlID,A.Name as Name,A.ID as ID,A.ID as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeAddon & " as NavIconType,A.Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
                End If
                If BuildVersion > "4.0.321" Then
                    SQL = "select" _
                        & " " & FieldList _
                        & " from ccAggregateFunctions A" _
                        & " left join ccAddonCollections C on C.ID=A.CollectionID" _
                        & " where C.ID is null" _
                        & " order by A.Name"
                Else
                    '
                    ' addoncollectionrules deprecated for collecitonid
                    '
                    SQL = "select" _
                        & " " & FieldList _
                        & " from ( ccAggregateFunctions A" _
                        & " left join ccAddonCollectionRules R on R.Addonid=A.id )" _
                        & " left join ccAddonCollections C on C.ID=R.CollectionID" _
                        & " where C.ID is null" _
                        & " order by A.Name"
                End If
                NodeType = NodeTypeAddon
                BlockSubNodes = True
                NodeIDString = ""
                CS = Main.OpenCSSQL("default", SQL)
                Do While Main.iscsok(CS)
                    Name = Trim(Main.getcsText(CS, "name"))
                    addonid = Main.GetCSInteger(CS, "AddonID")
                    NavIconTitleHtmlEncoded = kmaEncodeHTML(Name)
                    ContentControlID = Main.GetCSInteger(CS, "ContentControlID")
                    s = s & GetNavigatorNode(0, ContentControlID, 0, 0, 0, "", addonid, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeAddon, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeAddon, 0, False, OpenNodeList, NodeIDString, NodeNavigatorJS)
                    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    Call Main.NextCSRecord(CS)
                Loop
                Call Main.closeCs(CS)
            Case NodeIDLegacyMenu
                '
                ' Special Case: build old top menus under this Navigator entry
                '
                SQL = GetMenuSQL("(parentid=0)or(parentid is null)", LegacyMenuContentName)
                CS = Main.OpenCSSQL("default", SQL)
                BlockSubNodes = False
                If Not Main.iscsok(CS) Then
                    '
                    ' Empty list, add to EmptyNodeList
                    '
                    EmptyNodeList = EmptyNodeList & "," & TopParentNode
                End If
            Case NodeIDAllContentList
                '
                ' special case: all content
                '
                FieldList = "Name,ID,0 as AddonID,0 as NewWindow,ID as ContentID,'' as LinkPage," & NavIconTypeContent & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid,0 as collectionid"
                CS = Main.openCSContent("Content", , , , , , FieldList)
                NodeType = NodeTypeContent
                BlockSubNodes = True
            Case "0", ""
                '
                ' Navigator Entries, list home(s) plus all roots
                '
                NodeType = NodeTypeEntry
                BlockSubNodes = False
                Link = kmaEncodeHTML("http://" & Main.ServerDomain)
                s = s & CR & "<div class=ccNavLink><A href=""" & Link & """>" & IconPublicHome & "</A>&nbsp;<A href=""" & Link & """>Public Home</A></div>"
                Link = kmaEncodeHTML(Main.SiteProperty_AdminURL)
                s = s & CR & "<div class=ccNavLink><A href=""" & Link & """>" & IconAdminHome & "</A>&nbsp;<A href=""" & Link & """>Admin Home</A></div>"
                CS = Main.OpenCSSQL("default", GetMenuSQL("((Parentid=0)or(Parentid is null))", NavigatorContentName))
                Do While Main.iscsok(CS)
                    Name = Trim(Main.getcsText(CS, "name"))
                    NavigatorID = Main.GetCSInteger(CS, "ID")
                    NavIconTitleHtmlEncoded = kmaEncodeHTML(Name)
                    NodeIDString = CStr(NavigatorID)
                    If AutoManageAddons Then
                        '
                        ' special cases - root nodes that do not just deliver menu entries
                        '
                        Select Case LCase(Name)
                            Case "manage add-ons"
                                NodeIDString = NodeIDManageAddons
                            Case "settings"
                                NodeIDString = NodeIDSettings
                            Case "tools"
                                NodeIDString = NodeIDTools
                            Case "reports"
                                NodeIDString = NodeIDReports
                        End Select
                    End If
                    If Name = "Manage Content" Then
                        Name = Name
                    End If
                    s = s & GetNavigatorNode(0, 0, 0, 0, 0, "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, NavigatorID, NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEntry, 0, False, OpenNodeList, NodeIDString, NodeNavigatorJS)
                    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    Call Main.NextCSRecord(CS)
                Loop
                Call Main.closeCs(CS)
                '
                ' Add a Legacy Menu node to the top-most parent menu at the very end
                '
                If kmaEncodeBoolean(Main.GetSiteProperty("AllowNavigatorLegacyEntry", "0")) Then
                    Name = "Legacy Menu"
                    NavIconTitleHtmlEncoded = "Legacy Menu"
                    NodeIDString = NodeIDLegacyMenu
                    s = s & GetNavigatorNode(0, 0, 0, 0, 0, "", 0, 0, Name, LegacyMenuControlID, EmptyNodeList, 0, NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEntry, False, BlockSubNodes, OpenNodeList, NodeIDString, NodeNavigatorJS)
                    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                End If
            Case NodeIDSettings
                '
                ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                '
                s = s & GetNavigator_MixedNode(EmptyNodeList, TopParentNode, LegacyMenuControlID, AutoManageAddons, NodeType, OpenNodeList, 3, Main.GetRecordID("navigator entries", "settings"), NavIconTypeSetting, NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
            Case NodeIDTools
                '
                ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                '
                s = s & GetNavigator_MixedNode(EmptyNodeList, TopParentNode, LegacyMenuControlID, AutoManageAddons, NodeType, OpenNodeList, 4, Main.GetRecordID("navigator entries", "tools"), NavIconTypeTool, NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
            Case NodeIDReports
                '
                ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                '
                s = s & GetNavigator_MixedNode(EmptyNodeList, TopParentNode, LegacyMenuControlID, AutoManageAddons, NodeType, OpenNodeList, 2, Main.GetRecordID("navigator entries", "reports"), NavIconTypeReport, NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
            Case Else
                '
                ' numeric node (default case) - list navigator records with parent=TopParentNode
                '
                CS = -1
                If IsNumeric(TopParentNode) Then
                    If InStr(1, EmptyNodeList & ",", "," & TopParentNode & ",") <> 0 Then
                        EmptyNodeList = EmptyNodeList
                    Else
                        '
                        ' Navigator Entries, child under TopParentNode
                        '
                        SQL = GetMenuSQL("parentid=" & TopParentNode, "")
                        CS = Main.OpenCSSQL("default", SQL)
                        BlockSubNodes = False
                        'NodeIDStringPrefix = "n"
                        If Not Main.iscsok(CS) Then
                            '
                            ' Empty list, add to EmptyNodeList
                            '
                            EmptyNodeList = EmptyNodeList & "," & TopParentNode
                            'Call Main.SaveBake(BakeName, EmptyNodeList, "Navigator Entries")
                        End If
                    End If
                End If
        End Select
        '
        ' ----- List Navigator Nodes, if not already displayed
        '
        If (Not Main.iscsok(CS)) And (NodeType = NodeTypeEntry) Then
            '
            ' No child nodes, if this node includes a CID, list the first 20 content records with a 'more'
            '
            ContentID = 0
            If IsNumeric(TopParentNode) Then
                Call Main.closeCs(CS)
                CS = Main.openCSContent(NavigatorContentName, "id=" & TopParentNode)
                If Main.iscsok(CS) Then
                    ContentID = Main.GetCSInteger(CS, "ContentID")
                End If
                If ContentID <> 0 Then
                    ContentID = ContentID
                    ContentName = Main.GetContentNameByID(ContentID)
                    If ContentName <> "" Then
                        'ContentTableName = Main.GetContentTablename(ContentName)
                        Call Main.closeCs(CS)
                        CS = Main.openCSContent(ContentName, , , , , , "ID,Name,ContentControlID", 20, 1)
                       ' SQL = "select top 20 " _
                       '     & " ID,Name,ContentControlID from " & ContentTableName
                       ' CS = Main.OpenCSSQL("default", SQL)
                        Ptr = 0
                        If Main.iscsok(CS) Then
                            EmptyNodeList = Replace(EmptyNodeList, "," & TopParentNode, "")
                            Do While Main.iscsok(CS) And Ptr < 20
                                NavigatorID = Main.GetCSInteger(CS, "ID")
                                RecordName = Main.getcsText(CS, "Name")
                                If RecordName = "" Then
                                    RecordName = "Record " & NavigatorID
                                End If
                                '
                                If Len(RecordName) > 53 Then
                                    RecordName = Left(RecordName, 25) & "..." & Right(RecordName, 25)
                                End If
                                NavIconTitleHtmlEncoded = kmaEncodeHTML("Edit '" & RecordName & "' in '" & ContentName & "'")
                                IconNoSubNodes = IconRecord
                                IconNoSubNodes = Replace(IconNoSubNodes, "{title}", NavIconTitleHtmlEncoded)
                                Link = "?id=" & NavigatorID & "&cid=" & Main.GetCSInteger(CS, "ContentControlID") & "&af=4"
                                ATag = "<a href=""" & Link & """ title=""" & NavIconTitleHtmlEncoded & """>"
                                s = s & CR & "<div class=""ccNavLink ccNavLinkEmpty"">" & ATag & IconNoSubNodes & "</a>&nbsp;" & ATag & RecordName & "</a></div>"
                                Ptr = Ptr + 1
                                Call Main.NextCSRecord(CS)
                            Loop
                            If Ptr = 20 Then
                                NavIconTitleHtmlEncoded = kmaEncodeHTML("Open All '" & NavigatorContentName & "'")
                                Link = "?cid=" & ContentID
                                s = s & CR & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconClosed & "&nbsp;<a href=""" & Link & """ title=""" & NavIconTitleHtmlEncoded & """>more...</a></div>"
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf Main.iscsok(CS) Then
            '
            ' List out child menus
            '
            Do While Main.iscsok(CS)
                CollectionID = Main.GetCSInteger(CS, "CollectionID")
                NavigatorID = Main.GetCSInteger(CS, "ID")
                Name = Trim(Main.getcsText(CS, "name"))
                NewWindow = Main.GetCSBoolean(CS, "newwindow")
                ContentID = Main.GetCSInteger(CS, "ContentID")
                Link = Trim(Main.getcsText(CS, "LinkPage"))
                addonid = Main.GetCSInteger(CS, "AddonID")
                NavIconType = Main.GetCSInteger(CS, "NavIconType")
                NavIconTitle = Main.getcsText(CS, "NavIconTitle")
                HelpAddonID = Main.GetCSInteger(CS, "HelpAddonID")
                If HelpAddonID <> 0 Then
                    HelpAddonID = HelpAddonID
                End If
                helpCollectionID = Main.GetCSInteger(CS, "HelpCollectionID")
                If NavIconTitle = "" Then
                    NavIconTitle = Name
                End If
                ContentControlID = Main.GetCSInteger(CS, "ContentControlID")
                If LCase(Name) = "all content" Then
                    '
                    ' special case: any Navigator Entry named 'all content' returns the content list
                    '
                    NodeIDString = NodeIDAllContentList
                Else
                    NodeIDString = CStr(NavigatorID)
                End If
                NavIconTitleHtmlEncoded = kmaEncodeHTML(NavIconTitle)
                s = s & GetNavigatorNode(CollectionID, ContentControlID, helpCollectionID, HelpAddonID, ContentID, Link, addonid, SettingPageID, Name, LegacyMenuControlID, EmptyNodeList, NavigatorID, NavIconType, NavIconTitleHtmlEncoded, AutoManageAddons, NodeType, NewWindow, BlockSubNodes, OpenNodeList, NodeIDString, NodeNavigatorJS)
                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                Call Main.NextCSRecord(CS)
            Loop
            Call Main.closeCs(CS)
        End If
        '
        '
        '
        If EmptyNodeListInitial <> EmptyNodeList Then
            Call Main.SaveBake(BakeName, EmptyNodeList, "Navigator Entries")
        End If
        GetNavigator = s
'        If Return_NavigatorJS <> "" Then
'            Return_NavigatorJS = "" _
'                & "if(window.navDrop) {" _
'                & Return_NavigatorJS _
'                & "};"
'            Call Main.AddHeadScriptCode(Return_NavigatorJS, "Admin Navigator")
'            'Call Main.AddHeadScriptCode("alert('addheadscriptcode');", "Admin Navigator")
'        End If
    End If
    '
    ' ----- Timer Trace
    'Main.ToolsPanelTimerTrace = Replace(Main.ToolsPanelTimerTrace, TimerTraceMarker, (GetTickCount - StartTickCount)) & CR & "</ul>"
    ' ----- /Timer Trace
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassTrapError("GetNavigator")
End Function
'
'========================================================================
'   list mixed nodes (settings/reports/tools)
'========================================================================
'

Friend Function GetNavigator_MixedNode(EmptyNodeList As String, TopParentNode As String, LegacyMenuControlID As Long, AutoManageAddons As Boolean, NodeType As NodeTypeEnum, OpenNodeList As String, AddonNavTypeID As Long, MenuParentNodeID As Long, AdminNavIconTypeSetting As Long, Return_DraggableJS As String) As String
    On Error GoTo ErrorTrap
    '
    Dim NodeDraggableJS As String
    Dim LastName As String
    Dim Index As New FastIndex4.FastIndexClass
    Dim SortNodes() As SortNodeType
    Dim SortPtr As Long
    Dim SortCnt As Long
    Dim test As String
    Dim NameSuffix As String
    Dim NodeID As Long
    Dim NodeIDString As String
    Dim SQL As String
    Dim CS As Long
    Dim Criteria As String
    Dim BlockSubNodes As Boolean
    '
    ' list mixed nodes (settings/reports/tools), includes menu nodes and addons with type='setting' sorted in
    '
    If InStr(1, EmptyNodeList & ",", "," & TopParentNode & ",") <> 0 Then
        EmptyNodeList = EmptyNodeList
    Else
        '
        ' Add addons to node list
        '
        NodeIDString = ""
        Criteria = "(navtypeid=" & AddonNavTypeID & ")"
        If (AddonNavTypeID = 2) Or (AddonNavTypeID = 3) Or (AddonNavTypeID = 4) Then
            '
            ' if setting, report or tool, "admin" is not needed
            '
        Else
            '
            ' for Manage Addons node, admin is needed for non-developer
            '
            If Not Main.isDeveloper() Then
                Criteria = Criteria & "and(admin<>0)"
            End If
        End If
        CS = Main.openCSContent("Add-ons", Criteria, , , , , "*")
        SortPtr = 0
        Do While Main.iscsok(CS)
            ReDim Preserve SortNodes(SortPtr)
            With SortNodes(SortPtr)
                .Name = Trim(Main.getcsText(CS, "name"))
                .addonid = Main.GetCSInteger(CS, "ID")
                .NavIconTitle = .Name
                .ContentControlID = Main.GetCSInteger(CS, "ContentControlID")
                .NavIconType = AdminNavIconTypeSetting
                Call Index.SetPointer(.Name, SortPtr)
            End With
            SortPtr = SortPtr + 1
            Call Main.NextCSRecord(CS)
        Loop
        Call Main.closeCs(CS)
        '
        ' Add real navigator nodes to node list
        '
        SQL = GetMenuSQL("parentid=" & MenuParentNodeID, "")
        CS = Main.OpenCSSQL("default", SQL)
        BlockSubNodes = False
        Do While Main.iscsok(CS)
            ReDim Preserve SortNodes(SortPtr)
            With SortNodes(SortPtr)
                .Name = Trim(Main.getcsText(CS, "name"))
                .NavigatorID = Main.GetCSInteger(CS, "ID")
                .CollectionID = Main.GetCSInteger(CS, "CollectionID")
                .NewWindow = Main.GetCSBoolean(CS, "newwindow")
                .ContentID = Main.GetCSInteger(CS, "ContentID")
                .Link = Trim(Main.getcsText(CS, "LinkPage"))
                .addonid = Main.GetCSInteger(CS, "AddonID")
                .NavIconType = Main.GetCSInteger(CS, "NavIconType")
                .NavIconTitle = Main.getcsText(CS, "NavIconTitle")
                .HelpAddonID = Main.GetCSInteger(CS, "HelpAddonID")
                .helpCollectionID = Main.GetCSInteger(CS, "HelpCollectionID")
                If .NavIconTitle = "" Then
                    .NavIconTitle = .Name
                End If
                .ContentControlID = Main.GetCSInteger(CS, "ContentControlID")
                If LCase(.Name) = "all content" Then
                    '
                    ' special case: any Navigator Entry named 'all content' returns the content list
                    '
                    .NodeIDString = NodeIDAllContentList
                Else
                    .NodeIDString = CStr(.NavigatorID)
                    .NavIconTitle = .Name
                    .ContentControlID = Main.GetCSInteger(CS, "ContentControlID")
                End If
                Call Index.SetPointer(.Name, SortPtr)
            End With
            SortPtr = SortPtr + 1
            Call Main.NextCSRecord(CS)
        Loop
        Call Main.closeCs(CS)
        test = Index.GetFirstPointer()
        If test = "" Then
            '
            ' Empty list, add to EmptyNodeList
            '
            EmptyNodeList = EmptyNodeList & "," & TopParentNode
        Else
            LastName = ""
            Do While test <> ""
                SortPtr = kmaEncodeInteger(test)
                With SortNodes(SortPtr)
                    If LCase(.Name) <> LastName Then
                        GetNavigator_MixedNode = GetNavigator_MixedNode & GetNavigatorNode(.CollectionID, .ContentControlID, .helpCollectionID, .HelpAddonID, .ContentID, .Link, .addonid, 0, .Name, LegacyMenuControlID, EmptyNodeList, .NavigatorID, .NavIconType, kmaEncodeHTML(.NavIconTitle), AutoManageAddons, NodeType, .NewWindow, BlockSubNodes, OpenNodeList, .NodeIDString, NodeDraggableJS)
                        Return_DraggableJS = Return_DraggableJS & NodeDraggableJS
                        LastName = LCase(.Name)
                    End If
                End With
                test = Index.GetNextPointer()
            Loop
        End If
    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassTrapError("GetNavigator_MixedNode", "Trap")
End Function


'
'
'
Private Function GetNavigatorNode(CollectionID As Long, ContentControlID As Long, helpCollectionID As Long, HelpAddonID As Long, ContentID As Long, Link As String, addonid As Long, ignore As Long, Name As String, LegacyMenuControlID As Long, EmptyNodeList As String, NavigatorID As Long, NavIconType As String, NavIconTitleHtmlEncoded As String, AutoManageAddons As Boolean, NodeType As NodeTypeEnum, NewWindow As Boolean, BlockSubNodes As Boolean, OpenNodeList As String, NodeIDString As String, Return_NavigatorJS As String) As String
    On Error GoTo ErrorTrap
    '
    Dim csCollection As Long
    Dim collectionName As String
    Dim collectionHelpLink As String
    Dim collectionHelp As String
    Dim NodeNavigatorJS As String
    Dim NavLinkHTMLId As String
    Dim SubNav As String
    Dim DivIDClosed As String
    Dim DivIDOpened As String
    Dim DivIDContent As String
    Dim DivIDEmpty As String
    Dim s As String
    Dim CSAddon As Long
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
    ' Determine is it has either a function, or child entries
    '
    BlockNode = False
    Return_NavigatorJS = ""
    WorkingName = Name
    IsVisible = (CollectionID <> 0) Or (ContentControlID = LegacyMenuControlID) Or (helpCollectionID <> 0) Or (HelpAddonID <> 0) Or (ContentID <> 0) Or (Link <> "") Or (addonid <> 0) Or (LCase(WorkingName) = "all content") Or (LCase(WorkingName) = "add-ons with no collection")
    If Not IsVisible Then
        '
        ' IsVisible if it is not in the EmptyNodeList (has child entries)
        '
        IsVisible = (InStr(1, "," & EmptyNodeList & ",", "," & NavigatorID & ",") = 0)
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
        Select Case LCase(NavIconType)
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
                        Case NodeTypeAddon
                            '
                            ' List of Addons
                            '
                            'NodeIDString = ""
                            DivIDBase = "a" & NavigatorID
                        Case NodeTypeCollection
                            '
                            ' List of Collections
                            '
                            'NodeIDString = "collection." & CollectionID
                            DivIDBase = "c" & CollectionID
                        Case NodeTypeContent
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
        '
        ' setup link
        '
        If Link <> "" Then
            NavLinkHTMLId = "n" & NavigatorID
            workingNameHtmlEncoded = kmaEncodeHTML(WorkingName)
            If NewWindow Then
                workingNameHtmlEncoded = "<a name=""navLink"" id=""" & NavLinkHTMLId & """ href=""" & Link & """ target=""_blank"" title=""Open '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
            Else
                workingNameHtmlEncoded = "<a name=""navLink"" id=""" & NavLinkHTMLId & """ href=""" & Link & """ title=""Open '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
            End If
        Else
            '
            ' If link page, use this
            '
            If addonid <> 0 Then
                '
                ' link to addon
                '
                Link = ""
                CSAddon = Main.OpenCSContentRecord("Add-ons", addonid, , , "remotemethod,name,ccguid")
                If Main.iscsok(CSAddon) Then
                    AddonGuid = Main.getcsText(CSAddon, "ccguid")
                    AddonName = Main.getcsText(CSAddon, "name")
                    If Main.GetCSBoolean(CSAddon, "remotemethod") Then
                        NewWindow = True
                        Link = Main.SiteProperty_AdminURL & "?" & RequestNameRemoteMethodAddon & "=" & kmaEncodeRequestVariable(AddonName)
                    End If
                End If
                Call Main.closeCs(CSAddon)
                If Link = "" Then
                    If AddonGuid <> "" Then
                        Link = kmaModifyLinkQuery(Main.SiteProperty_AdminURL, "addonguid", AddonGuid, True)
                    Else
                        Link = kmaModifyLinkQuery(Main.SiteProperty_AdminURL, "addonid", CStr(addonid), True)
                    End If
                End If
                NavLinkHTMLId = "a" & addonid
                workingNameHtmlEncoded = kmaEncodeHTML(WorkingName)
                If NewWindow Then
                    workingNameHtmlEncoded = "<a name=""navLink"" id=""" & NavLinkHTMLId & """ href=""" & Link & """ target=""_blank"" title=""Run '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                Else
                    workingNameHtmlEncoded = "<a name=""navLink"" id=""" & NavLinkHTMLId & """ href=""" & Link & """ title=""Run '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                End If
            ElseIf ContentID <> 0 Then
                '
                ' go edit the content
                '
                Link = kmaModifyLinkQuery(Main.SiteProperty_AdminURL, "cid", CStr(ContentID), True)
                NavLinkHTMLId = "c" & ContentID
                workingNameHtmlEncoded = kmaEncodeHTML(WorkingName)
                If NewWindow Then
                    workingNameHtmlEncoded = "<a name=""navLink"" id=""" & NavLinkHTMLId & """ href=""" & Link & """ target=""_blank"" title=""List All '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                Else
                    workingNameHtmlEncoded = "<a name=""navLink"" id=""" & NavLinkHTMLId & """ href=""" & Link & """ title=""List All '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                End If
            ElseIf HelpAddonID <> 0 Then
                '
                ' go to Addon Help
                '
                Link = kmaModifyLinkQuery(Main.SiteProperty_AdminURL, "helpaddonid", CStr(HelpAddonID), True)
                workingNameHtmlEncoded = kmaEncodeHTML(WorkingName)
                If NewWindow Then
                    workingNameHtmlEncoded = "<a href=""" & Link & """ target=""_blank"" title=""Help for Add-on '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                Else
                    workingNameHtmlEncoded = "<a href=""" & Link & """ title=""Help for Add-on '" & NavIconTitleHtmlEncoded & "'"">" & NavIconTitleHtmlEncoded & "</a>"
                End If
            ElseIf helpCollectionID <> 0 Then
                '
                ' go to Collection Help
                '
                csCollection = Main.openCSContent("add-on collections", "id=" & helpCollectionID, , , , , "name,helpLink,help")
                If Not Main.iscsok(csCollection) Then
                    BlockNode = True
                Else
                    collectionName = Main.getcsText(csCollection, "name")
                    collectionHelpLink = Main.getcsText(csCollection, "helpLink")
                    collectionHelp = Main.getcsText(csCollection, "help")
                    '
                    WorkingName = collectionName
                    Link = collectionHelpLink
                    workingNameHtmlEncoded = kmaEncodeHTML(WorkingName)
                    If Link <> "" Then
                        NewWindow = True
                    ElseIf (collectionHelp <> "") Then
                        Link = kmaModifyLinkQuery(Main.SiteProperty_AdminURL, "helpcollectionid", CStr(helpCollectionID), True)
                    Else
                        BlockNode = True
                    End If
                    If Not BlockNode Then
                        If NewWindow Then
                            workingNameHtmlEncoded = "<a href=""" & Link & """ target=""_blank"" title=""Help for Collection '" & workingNameHtmlEncoded & "'"">Help</a>"
                        Else
                            workingNameHtmlEncoded = "<a href=""" & Link & """ title=""Help for Collection '" & workingNameHtmlEncoded & "'"">Help</a>"
                        End If
                    End If
                End If
                Call Main.closeCs(csCollection)
            Else
                workingNameHtmlEncoded = kmaEncodeHTML(WorkingName)
            End If
        End If
        '
        If Not BlockNode Then
            If BlockSubNodes Then
                '
                ' This is a hardcoded item (like Add-on), it has no subnodes
                '
                s = s & CR & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & "</div>"
            Else
                '
                DivIDClosed = DivIDBase & "a"
                DivIDOpened = DivIDBase & "b"
                DivIDContent = DivIDBase & "c"
                DivIDEmpty = DivIDBase & "d"
                
                If (ContentID = 0) And (InStr(1, EmptyNodeList & ",", "," & NodeIDString & ",") <> 0) Then
                    '
                    ' In EmptyNodeList
                    '
                    s = s & CR & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & "&nbsp;" & workingNameHtmlEncoded & "</div>"
                ElseIf InStr(1, OpenNodeList & ",", "," & NodeIDString & ",") <> 0 Then
                    '
                    ' This node is open
                    '
                    SubNav = GetNavigator(NodeIDString, OpenNodeList, NodeNavigatorJS)
                    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    If SubNav <> "" Then
                        '
                        ' display the subnav
                        '
                        s = s & CR & "<div class=ccNavLink ID=" & DivIDClosed & " style=""display:none;""><A class=""ccNavClosed"" href=""#"" onclick=""AdminNavOpenClick('" & DivIDClosed & "','" & DivIDOpened & "','" & DivIDContent & "','" & NodeIDString & "','" & DivIDEmpty & "');return false;"">" & IconClosed & "</A>&nbsp;" & workingNameHtmlEncoded & "</div>"
                        s = s & CR & "<div class=ccNavLink ID=" & DivIDOpened & "><A class=""ccNavOpened"" href=""#"" onclick=""AdminNavCloseClick('" & DivIDOpened & "','" & DivIDClosed & "','" & DivIDContent & "','" & NodeIDString & "');return false;"">" & IconOpened & "</A>&nbsp;" & workingNameHtmlEncoded & "</div>"
                        s = s _
                            & CR & "<div class=ccNavLinkChild ID=" & DivIDContent & ">" _
                            & kmaIndent(SubNav) _
                            & CR & "</div>"
                    Else
                        '
                        ' it has a NO subnav
                        '
                        s = s & CR & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & "&nbsp;" & workingNameHtmlEncoded & "</div>"
                    End If
                Else
                    '
                    ' This node is closed
                    '
                    s = s & CR & "<div class=ccNavLink ID=" & DivIDClosed & " ><A class=""ccNavClosed"" href=""#"" onclick=""AdminNavOpenClick('" & DivIDClosed & "','" & DivIDOpened & "','" & DivIDContent & "','" & NodeIDString & "','','" & DivIDContent & "');return false;"">" & IconClosed & "</A>&nbsp;" & workingNameHtmlEncoded & "</div>"
                    s = s & CR & "<div class=ccNavLink ID=" & DivIDOpened & " style=""display:none;""><A class=""ccNavOpened"" href=""#"" onclick=""AdminNavCloseClick('" & DivIDOpened & "','" & DivIDClosed & "','" & DivIDContent & "','" & NodeIDString & "');return false;"">" & IconOpened & "</A>&nbsp;" & workingNameHtmlEncoded & "</div>"
                    s = s & CR & "<div class=""ccNavLink ccNavLinkEmpty"" ID=" & DivIDEmpty & " style=""display:none;"">" & IconNoSubNodes & "&nbsp;" & workingNameHtmlEncoded & "</div>"
                    s = s & CR & "<div class=ccNavLinkChild ID=" & DivIDContent & " style=""display:none;margin-left:20px;"">&nbsp;&nbsp;&nbsp;&nbsp;<img src=""/cclib/images/ajax-loader-small.gif"" width=""16"" height=""16""></div>"
                End If
            End If
            If NavLinkHTMLId <> "" Then
                Return_NavigatorJS = Return_NavigatorJS _
                & CR & "$(function(){" _
                    & "$('#" & NavLinkHTMLId & "').draggable({" _
                        & "opacity: 0.50" _
                        & ",helper: 'clone'" _
                        & ",revert: 'invalid'" _
                        & ",stop: function(event, ui){" _
                            & "navDrop('" & NavLinkHTMLId & "',ui.position.left,ui.position.top);" _
                        & "}" _
                        & ",cursor: 'move'" _
                    & "});" _
                & "});"
            End If
        End If
    End If
    '
    GetNavigatorNode = s
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassTrapError("GetNavigatorNode", "Trap")
End Function
'
'========================================================================
' Get sql for menu
'   if MenuContentName is blank, it will select values from either cdef
'========================================================================
'
Private Function GetMenuSQL(ParentCriteria As String, MenuContentName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CMCriteria As String
    Dim iParentCriteria As String
    Dim Criteria As String
    Dim SQL As String
    Dim ContentControlCriteria As String
    Dim SelectList As String
    Dim ContentManagementList As String
    '
    Criteria = "(Active<>0)"
    If MenuContentName <> "" Then
        Criteria = Criteria & "AND" & Main.GetContentControlCriteria(MenuContentName)
    End If
    iParentCriteria = KmaEncodeMissingText(ParentCriteria, "")
    If Main.isDeveloper Then
        '
        ' ----- Developer
        '
    ElseIf Main.IsAdmin Then
        '
        ' ----- Administrator
        '
        Criteria = Criteria _
            & "AND((DeveloperOnly is null)or(DeveloperOnly=0))" _
            & "AND(ID in (" _
            & " SELECT AllowedEntries.ID" _
            & " FROM CCMenuEntries AllowedEntries LEFT JOIN ccContent ON AllowedEntries.ContentID = ccContent.ID" _
            & " Where ((ccContent.Active<>0)And((ccContent.DeveloperOnly is null)or(ccContent.DeveloperOnly=0)))" _
                & "OR(ccContent.ID Is Null)" _
            & "))"
    Else
        '
        ' ----- Content Manager
        '
        ContentManagementList = GetContentManagementList
        If ContentManagementList = "" Then
            CMCriteria = "(1=0)"
        Else
            'ContentManagementList = Mid(ContentManagementList, 2, Len(ContentManagementList) - 2)
            If InStr(1, ContentManagementList, ",") = 0 Then
                CMCriteria = "(ccContent.ID=" & ContentManagementList & ")"
            Else
                CMCriteria = "(ccContent.ID in (" & ContentManagementList & "))"
            End If
        End If
        
        Criteria = Criteria _
            & "AND((DeveloperOnly is null)or(DeveloperOnly=0))" _
            & "AND((AdminOnly is null)or(AdminOnly=0))" _
            & "AND(ID in (" _
            & " SELECT AllowedEntries.ID" _
            & " FROM CCMenuEntries AllowedEntries LEFT JOIN ccContent ON AllowedEntries.ContentID = ccContent.ID" _
            & " Where (" & CMCriteria & "and(ccContent.Active<>0)And((ccContent.DeveloperOnly is null)or(ccContent.DeveloperOnly=0))And((ccContent.AdminOnly is null)or(ccContent.AdminOnly=0)))" _
                & "OR(ccContent.ID Is Null)" _
            & "))"
        End If
    If iParentCriteria <> "" Then
        Criteria = "(" & iParentCriteria & ")AND" & Criteria
    End If
    If Main.SiteProperty_BuildVersion > "3.4.044" Then
        SelectList = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, ccMenuEntries.AddonID, ccMenuEntries.NavIconType, ccMenuEntries.NavIconTitle, HelpAddonID,HelpCollectionID,0 as collectionid"
    ElseIf Main.SiteProperty_BuildVersion > "3.3.982" Then
        SelectList = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, ccMenuEntries.AddonID, ccMenuEntries.NavIconType, ccMenuEntries.NavIconTitle,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
    ElseIf Main.SiteProperty_BuildVersion > "3.3.865" Then
        SelectList = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, ccMenuEntries.AddonID, ccMenuEntries.NavIconType, ccMenuEntries.NavIconTitle,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
    ElseIf Main.SiteProperty_BuildVersion > "3.3.686" Then
        SelectList = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, ccMenuEntries.AddonID, 'folder' as NavIconType, ccMenuEntries.Name as NavIconTitle,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
    Else
        SelectList = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, 0 as AddonID, 'folder' as NavIconType, ccMenuEntries.Name as NavIconTitle,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
    End If
    GetMenuSQL = "select " & SelectList & " from ccMenuEntries where " & Criteria & " order by ccMenuEntries.Name"
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassTrapError("GetMenuSQL", "Trap")
    '
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
'
'===========================================================================
'   Get Authoring List
'       returns a comma delimited list of ContentIDs that the Member can author
'===========================================================================
'
Private Function GetContentManagementList() As String
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim RS As Recordset
    Dim CIDArray() As Variant
    Dim CIDCount As Long
    Dim CIDPointer As Long
    'Dim cdef As CDefType
    Dim ContentName As String
    Dim ContentID As Long
    Dim ChildIDList As String
    Dim sqlTablememberRules As String
    '
    sqlTablememberRules = Main.GetContentTablename("Member Rules")
    '
    SQL = "Select ccGroupRules.ContentID as ID" _
        & " FROM ((" & sqlTablememberRules _
        & " Left Join ccGroupRules on " & sqlTablememberRules & ".GroupID=ccGroupRules.GroupID)" _
        & " Left Join ccContent on ccGroupRules.ContentID=ccContent.ID)" _
        & " WHERE" _
            & " (" & sqlTablememberRules & ".MemberID=" & Main.memberID & ")" _
            & " AND(ccGroupRules.Active<>0)" _
            & " AND(ccContent.Active<>0)" _
            & " AND(" & sqlTablememberRules & ".Active<>0)"
    Set RS = Main.openRSSQL("Default", SQL)
    If IsRSOK(RS) Then
        CIDArray = RS.GetRows()
        CIDCount = UBound(CIDArray, 2) + 1
    End If
    Set RS = Nothing
    For CIDPointer = 0 To CIDCount - 1
        ContentID = kmaEncodeInteger(CIDArray(0, CIDPointer))
        GetContentManagementList = GetContentManagementList & "," & CStr(ContentID)
        ContentName = Main.GetContentNameByID(ContentID)
        If ContentName <> "" Then
            ChildIDList = Main.GetContentProperty(ContentName, "ChildIDList")
            'cdef = Main.GetContentDefinition(ContentName)
            If ChildIDList <> "" Then
                GetContentManagementList = GetContentManagementList & "," & ChildIDList
            End If
        End If
    Next
    If Len(GetContentManagementList) > 1 Then
        GetContentManagementList = Mid(GetContentManagementList, 2)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetContentManagementList", "")
End Function



Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.adminNavigator
    Public Class getNodeClass
        Inherits AddonBaseClass
        '
        Public Structure NavigatorEnvironment
            Public adminUrl As String
            Public isDeveloper As Boolean
            Public buildVersion As String
            Public addonEditAddonUrlPrefix As String
            Public addonEditCollectionUrlPrefix As String
            Public contentFieldEditToolPrefix As String
            Public cacheDependencyList As String
        End Structure
        ''' <summary>
        ''' Return the full navigator
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                '
                ' setup environment
                Dim env As New NavigatorEnvironment
                env.adminUrl = CP.Site.GetText("adminurl")
                env.buildVersion = CP.Site.GetProperty("buildversion")
                env.isDeveloper = CP.User.IsDeveloper()
                If env.isDeveloper Then
                    env.addonEditAddonUrlPrefix = env.adminUrl & "?cid=" & CP.Content.GetID("add-ons") & "&af=4&id="
                    env.addonEditCollectionUrlPrefix = env.adminUrl & "?cid=" & CP.Content.GetID("add-on collections") & "&af=4&id="
                    env.contentFieldEditToolPrefix = env.adminUrl & "?af=105&contentid="
                End If
                env.cacheDependencyList = C51CacheController.createDependencyKeyInvalidateOnChange(CP, "ccMenuEntries", "default") & "," & C51CacheController.createDependencyKeyInvalidateOnChange(CP, "ccaggregatefunctions", "default")
                '
                Dim OpenNodeList As String = CP.Visit.GetText("AdminNavOpenNodeList")
                Dim ParentNode As String = CP.Doc.GetText("nodeid")
                If (ParentNode <> "") Then
                    If OpenNodeList = "" Then
                        OpenNodeList = "," & ParentNode
                        Call CP.Visit.SetProperty("AdminNavOpenNodeList", OpenNodeList)
                    Else
                        If ((OpenNodeList & ",").IndexOf("," & ParentNode.ToString & ",") < 0) Then
                            OpenNodeList = OpenNodeList & "," & ParentNode
                            Call CP.Visit.SetProperty("AdminNavOpenNodeList", OpenNodeList)
                        End If
                    End If
                End If

                Dim NavigatorJS As String = ""
                returnHtml = GetNodeList(CP, env, ParentNode, OpenNodeList, NavigatorJS)
                If NavigatorJS <> "" Then
                    NavigatorJS = "/* block if no dashboard js */if(window.navDrop) {" & NavigatorJS & "};"
                    Call CP.Doc.AddBodyEnd("<script language=""javascript"" type=""text/javascript/"">jQuery(document).ready(function(){" & NavigatorJS & "})</script>")
                End If
            Catch ex As Exception
                HandleError(CP, ex)
            End Try
            Return returnHtml
        End Function
        ''' <summary>
        ''' get Node list?
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="env"></param>
        ''' <param name="ParentNode"></param>
        ''' <param name="OpenNodeList"></param>
        ''' <param name="Return_NavigatorJS"></param>
        ''' <returns></returns>
        Friend Function GetNodeList(cp As CPBaseClass, env As NavigatorEnvironment, ParentNode As String, OpenNodeList As String, ByRef Return_NavigatorJS As String) As String
            Dim returnNav As String = ""
            Try
                Const AutoManageAddons = True
                Dim Index As New keyPtrIndexClass
                Dim cacheKey As String
                If env.isDeveloper Then
                    cacheKey = "AdminNav EmptyNodeList Dev"
                ElseIf cp.User.IsAdmin() Then
                    cacheKey = "AdminNav EmptyNodeList Admin"
                Else
                    cacheKey = "AdminNav EmptyNodeList CM" & cp.User.Id
                End If
                Dim EmptyNodeList As String = cp.Cache.Read(cacheKey)
                Dim SQL As String
                If EmptyNodeList <> "" Then
                    Call cp.Site.TestPoint("adminNavigator, emptyNodeList from cache=[" & EmptyNodeList & "]")
                Else
                    SQL = "select n.ID from ccMenuEntries n left join ccMenuEntries c on c.parentid=n.id Where c.ID Is Null group by n.id"
                    Dim cs2 As CPCSBaseClass = cp.CSNew()
                    If cs2.OpenSQL(SQL) Then
                        Do
                            EmptyNodeList &= "," & cs2.GetText("id")
                            Call cs2.GoNext()
                        Loop While cs2.OK()
                        EmptyNodeList = EmptyNodeList.Substring(1)
                    End If
                    Call cs2.Close()
                    Call cp.Site.TestPoint("adminNavigator, emptyNodeList from db=[" & EmptyNodeList & "]")
                    Call cp.Cache.Save(cacheKey, EmptyNodeList, env.cacheDependencyList)
                End If
                Dim EmptyNodeListInitial As String = EmptyNodeList
                Dim TopParentNode As String = ParentNode
                Dim parentNodeStack() As String
                If TopParentNode = "" Then
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
                Dim NodeIDString As String
                Dim IconNoSubNodes As String
                Dim NavigatorID As Integer
                Dim CollectionID As Integer
                Dim s As String = ""
                Dim Name As String
                Dim ContentID As Integer
                Dim addonid As Integer
                Dim Link As String
                Dim Criteria As String
                Dim BlockSubNodes As Boolean
                Dim NodeType As NodeTypeEnum
                Dim NavIconType As Integer
                Dim NavIconTitle As String
                Dim NavIconTitleHtmlEncoded As String
                Dim ContentControlID As Integer
                Dim csChildList As CPCSBaseClass = cp.CSNew()
                Dim linkSuffixList As String
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
                        '
                        NodeIDString = ""
                        addonid = 0
                        Dim addonManagerVerionGuid As String = AddonManagerGuid
                        If (cp.Version.Substring(0, 1) = "4") Then
                            '
                            ' -- everything before v5
                            addonManagerVerionGuid = AddonManagerGuid
                            addonid = cp.Content.GetRecordID("Add-ons", "Add-on Manager")
                        ElseIf (cp.Version.Substring(0, 4) = "5.00") Then
                            '
                            ' -- version 5.00
                            Dim cs As CPCSBaseClass = cp.CSNew()
                            addonManagerVerionGuid = "{3d643ebd-c482-49ac-99ca-03aabb28ee83}"
                            If cs.Open("add-ons", "ccguid='{3d643ebd-c482-49ac-99ca-03aabb28ee83}'") Then
                                addonid = cs.GetInteger("id")
                            End If
                            cs.Close()
                        Else
                            '
                            ' -- version 5.01+
                            Dim cs As CPCSBaseClass = cp.CSNew()
                            addonManagerVerionGuid = "{4DD876E7-BCC4-4AF0-B32E-59FFAB816478}"
                            If cs.Open("add-ons", "ccguid='{4DD876E7-BCC4-4AF0-B32E-59FFAB816478}'") Then
                                addonid = cs.GetInteger("id")
                            End If
                            cs.Close()
                        End If
                        If (addonid > 0) Then
                            s = s & GetNode(cp, env, 0, 0, 0, 0, 0, "?addonguid=" & addonManagerVerionGuid, addonid, 0, "Add-on Manager", EmptyNodeList, 0, NavIconTypeAddon, "Add-on Manager", AutoManageAddons, NodeTypeEnum.NodeTypeAddon, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                        End If
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        '
                        ' List Collections
                        '
                        FieldList = "Name,0 as id,ccaddoncollections.id as collectionid,0 as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeFolder & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid,blockNavigatorNode,system"
                        'FieldList = "Name,id as collectionid,0 as ID,0 as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeFolder & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid"
                        Criteria = "((system=0)or(system is null))"
                        If Not env.isDeveloper Then
                            'Criteria = "((system=0)or(system is null))"
                            If (env.buildVersion >= "4.1.512") Then
                                Criteria = Criteria & "and((blockNavigatorNode=0)or(blockNavigatorNode is null))"
                            End If
                        End If
                        Dim cs3 As CPCSBaseClass = cp.CSNew()
                        NodeType = NodeTypeEnum.NodeTypeCollection
                        BlockSubNodes = False
                        If cs3.Open("Add-on Collections", Criteria, "name", True, FieldList, 9999, 1) Then
                            Do
                                Name = Trim(cs3.GetText("name"))
                                NavIconTitle = Name
                                CollectionID = cs3.GetInteger("collectionid")
                                NodeIDString = NodeIDManageAddonsCollectionPrefix & "." & CollectionID
                                NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle)
                                linkSuffixList = ""
                                If env.isDeveloper Then
                                    linkSuffixList &= "<a href=""" & env.addonEditCollectionUrlPrefix & CollectionID & """>edit</a>"
                                    If cs3.GetBoolean("system") Then
                                        linkSuffixList &= ",sys"
                                    End If
                                    If cs3.GetBoolean("blockNavigatorNode") Then
                                        linkSuffixList &= ",dev"
                                    End If
                                    linkSuffixList = "&nbsp;(" & linkSuffixList & ")"
                                End If
                                s = s & GetNode(cp, env, CollectionID, 0, 0, 0, 0, "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeAddon, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeCollection, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS, linkSuffixList)
                                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                                Call cs3.GoNext()
                            Loop While cs3.OK()
                        End If
                        Call cs3.Close()
                        'CS = Main.openCSContent("Add-on Collections", Criteria, , , , , FieldList)
                        'NodeType = NodeTypeCollection
                        'BlockSubNodes = False
                        'Do While Main.iscsok(CS)
                        '    Name = Trim(csx.getText( "name"))
                        '    NavIconTitle = Name
                        '    CollectionID = csx.getInteger( "collectionid")
                        '    NodeIDString = NodeIDManageAddonsCollectionPrefix & "." & CollectionID
                        '    NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle)
                        '    s = s & GetNavigatorNode(cp,CollectionID, 0, 0, 0, 0, "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeAddon, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeCollection, false, False, OpenNodeList, NodeIDString, NodeNavigatorJS,"")
                        '    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        '    Call Main.NextCSRecord(CS)
                        'Loop
                        'Call Main.closeCs(CS)
                        '
                        ' Advanced folder to contain edit links to create addons and collections
                        '
                        NodeIDString = NodeIDManageAddonsAdvanced
                        s = s & GetNode(cp, env, 0, 0, 0, 0, 0, "", 0, 0, "Advanced", EmptyNodeList, 0, NavIconTypeFolder, "Add-ons With No Collection", AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
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
                        nodeHtml = cp.Cache.Read(cacheName)
                        If nodeHtml = "" Then
                            '
                            ' Help Icon
                            '
                            Name = "Help"
                            NodeIDString = ""
                            NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle)
                            nodeHtml &= GetNode(cp, env, 0, 0, CollectionID, 0, 0, "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeHelp, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, True, OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                            Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                            '
                            ' List out add-ons in this collection
                            '
                            NodeIDString = ""
                            FieldList = "*"
                            'FieldList = "Name,id as collectionid,0 as ID,0 as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeFolder & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid"
                            Criteria = "(collectionid=" & CollectionID & ")"
                            If Not env.isDeveloper Then
                                Criteria = Criteria & "and(admin<>0)"
                                'Criteria = Criteria & "and((template<>0)or(page<>0)or(admin<>0))"
                            End If
                            Dim cs4 As CPCSBaseClass = cp.CSNew()
                            Dim NameSuffix As String
                            If cs4.Open("add-ons", Criteria, "name", True, FieldList) Then
                                Do
                                    Name = Trim(cs4.GetText("name"))
                                    NameSuffix = ""
                                    linkSuffixList = ""
                                    If env.isDeveloper Then
                                        linkSuffixList &= "<a href=""" & env.addonEditAddonUrlPrefix & cs4.GetInteger("id") & """>edit</a>"
                                        If Not cs4.GetBoolean("admin") Then
                                            linkSuffixList &= ",dev"
                                        End If
                                        linkSuffixList = "&nbsp;(" & linkSuffixList & ")"
                                        'If cs4.GetBoolean("content") Then
                                        '    NameSuffix = NameSuffix & "c"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        'If cs4.GetBoolean("template") Then
                                        '    NameSuffix = NameSuffix & "t"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        ''Name = Name & "(" & NameSuffix & ")"
                                        'If cs4.GetBoolean("email") Then
                                        '    NameSuffix = NameSuffix & "m"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        'If cs4.GetBoolean("admin") Then
                                        '    NameSuffix = NameSuffix & "n"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        'If cs4.GetBoolean("remotemethod") Then
                                        '    NameSuffix = NameSuffix & "r"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        'If cs4.GetBoolean("onpagestartevent") Then
                                        '    NameSuffix = NameSuffix & "b"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        'If cs4.GetBoolean("onpageendevent") Then
                                        '    NameSuffix = NameSuffix & "a"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        'If cs4.GetBoolean("onbodystart") Then
                                        '    NameSuffix = NameSuffix & "s"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        'If cs4.GetBoolean("onpageendevent") Then
                                        '    NameSuffix = NameSuffix & "e"
                                        'Else
                                        '    NameSuffix = NameSuffix & "-"
                                        'End If
                                        'Name = Name & " (" & NameSuffix & ")"
                                    End If
                                    addonid = cs4.GetInteger("ID")
                                    NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name)
                                    ContentControlID = cs4.GetInteger("ContentControlID")
                                    Select Case cs4.GetInteger("navtypeid")
                                        Case 2
                                            NavIconType = NavIconTypeReport
                                        Case 3
                                            NavIconType = NavIconTypeSetting
                                        Case 4
                                            NavIconType = NavIconTypeTool
                                        Case Else
                                            NavIconType = NavIconTypeAddon
                                    End Select
                                    nodeHtml &= GetNode(cp, env, 0, ContentControlID, 0, 0, 0, "", addonid, 0, Name, EmptyNodeList, 0, NavIconType, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeAddon, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS, linkSuffixList)
                                    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                                    Call cs4.GoNext()
                                Loop While cs4.OK
                            End If
                            Call cs4.Close()
                            '
                            ' List out cdefs connected to this collection
                            '
                            NodeIDString = ""
                            Criteria = "(collectionid=" & CollectionID & ")"
                            If env.isDeveloper Then
                            ElseIf cp.User.IsAdmin Then
                                Criteria = Criteria & "and(developeronly=0)"
                            Else
                                Criteria = Criteria & "and(developeronly=0)and(adminonly=0)"
                            End If
                            Dim LastContentID As Integer
                            Dim DupsFound As Boolean
                            LastContentID = -1
                            DupsFound = False
                            SQL = "select c.id,c.name,c.contentcontrolid,c.developeronly,c.adminonly from ccContent c left join ccAddonCollectionCDefRules r on r.contentid=c.id where " & Criteria & " order by c.name"
                            Dim cs7 As CPCSBaseClass = cp.CSNew()
                            If cs7.OpenSQL(SQL) Then
                                Do
                                    Name = Trim(cs7.GetText("name"))
                                    ContentID = cs7.GetInteger("id")
                                    If ContentID = LastContentID Then
                                        DupsFound = True
                                    Else
                                        linkSuffixList = ""
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
                                        nodeHtml &= GetNode(cp, env, 0, ContentControlID, 0, 0, ContentID, "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeContent, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeContent, False, True, OpenNodeList, NodeIDString, NodeNavigatorJS, linkSuffixList)
                                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
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
                            Dim sqlCriteria As String = ""

                            If cs7.Open("add-on collections", "id=" & CollectionID) Then
                                dataRecordList = cs7.GetText("dataRecordList")
                            End If
                            Call cs7.Close()
                            If Not String.IsNullOrEmpty(dataRecordList) Then
                                dataRecords.AddRange(dataRecordList.Split(vbCrLf.ToCharArray))
                                For Each dataRecord As String In dataRecords
                                    dataRecordParts = dataRecord.Split(",".ToCharArray)
                                    dataRecordCdefName = dataRecordParts(0)
                                    If Not String.IsNullOrEmpty(dataRecordCdefName) Then
                                        dataRecordCdefID = cp.Content.GetID(dataRecordCdefName)
                                        If dataRecordCdefID <> 0 Then
                                            sqlCriteria = ""
                                            If dataRecordParts.Length >= 2 Then
                                                ' 
                                                ' contentname,(id or guid)
                                                '
                                                If dataRecordParts(1).Substring(0, 1) = "{" Then
                                                    sqlCriteria = "ccguid=" & cp.Db.EncodeSQLText(dataRecordParts(1))
                                                Else
                                                    sqlCriteria = "name=" & cp.Db.EncodeSQLText(dataRecordParts(1))
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
                                                    'nodeHtml &= GetNode(cp, env, 0, 0, 0, 0, 0, "/admin?af=4&cid=" & dataRecordCdefID.ToString & "&id=" & cs7.GetInteger("id"), 0, 0, dataRecordCdefName & ":" & cs7.GetText("name"), EmptyNodeList, 0, NavIconTypeContent, "Data", AutoManageAddons, NodeTypeEnum.NodeTypeContent, False, True, OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                                                    Call cs7.GoNext()
                                                Loop While cs7.OK

                                            End If
                                            Call cs7.Close()
                                        End If
                                    End If
                                Next
                            End If
                            '
                            If DupsFound Then
                                SQL = "select b.id from ccAddonCollectionCDefRules a,ccAddonCollectionCDefRules b where (a.id<b.id) and (a.contentid=b.contentid) and (a.collectionid=b.collectionid)"
                                SQL = "delete from ccAddonCollectionCDefRules where id in (" & SQL & ")"
                                Call cp.Db.ExecuteSQL(SQL)
                            End If
                            Call cp.Cache.Save(cacheName, nodeHtml, env.cacheDependencyList, Now.AddHours(1))
                        End If
                        s &= nodeHtml
                    Case NodeIDManageAddonsAdvanced
                        '
                        ' Special Case: clicked on Manage Add-ons.advanced
                        '   edit links for Add-ons, Add-on Collections
                        '
                        ' Folder to Add-ons without Collections
                        '
                        NodeIDString = NodeIDAddonsNoCollection
                        s = s & GetNode(cp, env, 0, 0, 0, 0, 0, "", 0, 0, "Add-ons With No Collection", EmptyNodeList, 0, NavIconTypeAddon, "Add-ons With No Collection", AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        '
                        Name = "Add-ons"
                        s = s & GetNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, OpenNodeList, "", NodeNavigatorJS, "")
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        '
                        Name = "Add-on Collections"
                        s = s & GetNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, OpenNodeList, "", NodeNavigatorJS, "")
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        '
                        Name = "Aggregate Access"
                        s = s & GetNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, OpenNodeList, "", NodeNavigatorJS, "")
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        '
                        Name = "Scripting Modules"
                        s = s & GetNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeContent, Name, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, OpenNodeList, "", NodeNavigatorJS, "")
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    Case NodeIDAddonsNoCollection
                        '
                        ' special case: Add-on List that do not have collections
                        '
                        CollectionID = 0
                        FieldList = "0 as ContentControlID,A.Name as Name,A.ID as ID,A.ID as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeAddon & " as NavIconType,A.Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid"
                        SQL = "select" _
                            & " " & FieldList _
                            & " from ccAggregateFunctions A" _
                            & " left join ccAddonCollections C on C.ID=A.CollectionID" _
                            & " where C.ID is null" _
                            & " order by A.Name"
                        NodeType = NodeTypeEnum.NodeTypeAddon
                        BlockSubNodes = True
                        NodeIDString = ""
                        Dim cs5 As CPCSBaseClass = cp.CSNew()
                        If cs5.OpenSQL(SQL) Then
                            Do
                                Name = Trim(cs5.GetText("name"))
                                addonid = cs5.GetInteger("AddonID")
                                NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name)
                                ContentControlID = cs5.GetInteger("ContentControlID")
                                s = s & GetNode(cp, env, 0, ContentControlID, 0, 0, 0, "", addonid, 0, Name, EmptyNodeList, 0, NavIconTypeAddon, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeAddon, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                                Call cs5.GoNext()
                            Loop While cs5.OK()
                        End If
                        Call cs5.Close()
                            '
                            'CS = Main.OpenCSSQL("default", SQL)
                            'Do While Main.iscsok(CS)
                            '    Name = Trim(csx.getText("name"))
                            '    addonid = csx.getInteger( "AddonID")
                            '    NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name)
                            '    ContentControlID = csx.getInteger( "ContentControlID")
                            '    s = s & GetNavigatorNode(cp,0, ContentControlID, 0, 0, 0, "", addonid, 0, Name, EmptyNodeList, 0, NavIconTypeAddon, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeAddon, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS,"")
                            '    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                            '    Call Main.NextCSRecord(CS)
                            'Loop
                            'Call Main.closeCs(CS)
                    Case NodeIDLegacyMenu
                        '
                        ' Special Case: build old top menus under this Navigator entry
                        '
                        BlockSubNodes = False
                        SQL = GetMenuSQL(cp, "(parentid=0)or(parentid is null)", LegacyMenuContentName)
                        If Not csChildList.OpenSQL(SQL) Then
                            '
                            ' Empty list, add to EmptyNodeList
                            '
                            EmptyNodeList = EmptyNodeList & "," & TopParentNode
                        End If
                            ''
                            'CS = Main.OpenCSSQL("default", SQL)
                            'BlockSubNodes = False
                            'If Not Main.iscsok(CS) Then
                            '    '
                            '    ' Empty list, add to EmptyNodeList
                            '    '
                            '    EmptyNodeList = EmptyNodeList & "," & TopParentNode
                            'End I
                    Case NodeIDAllContentList
                        '
                        ' special case: all content
                        '
                        FieldList = "Name,ID,0 as AddonID,0 as NewWindow,ID as ContentID,'' as LinkPage," & NavIconTypeContent & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid,0 as collectionid"
                        SQL = "select " & FieldList & " from cccontent order by name"
                        csChildList.OpenSQL(SQL)
                        'Call csChildList.Open("content", , "name", , FieldList, 999)
                        'CS = Main.openCSContent("Content", , , , , , FieldList)
                        NodeType = NodeTypeEnum.NodeTypeContent
                        BlockSubNodes = True
                    Case "0", ""
                        '
                        ' Navigator Entries, list home(s) plus all roots
                        '
                        NodeType = NodeTypeEnum.NodeTypeEntry
                        BlockSubNodes = False
                        Link = cp.Utils.EncodeHTML("http://" & cp.Site.Domain)
                        s = s & cr & "<div class=ccNavLink><A href=""" & Link & """>" & IconPublicHome & "</A>&nbsp;<A href=""" & Link & """>Public Home</A></div>"
                        Link = cp.Utils.EncodeHTML(cp.Site.GetText("adminUrl"))
                        s = s & cr & "<div class=ccNavLink><A href=""" & Link & """>" & IconAdminHome & "</A>&nbsp;<A href=""" & Link & """>Admin Home</A></div>"
                        Dim cs8 As CPCSBaseClass = cp.CSNew()
                        If cs8.OpenSQL(GetMenuSQL(cp, "((Parentid=0)or(Parentid is null))", NavigatorContentName)) Then
                            Do
                                Name = Trim(cs8.GetText("name"))
                                NavigatorID = cs8.GetInteger("ID")
                                NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name)
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
                                s = s & GetNode(cp, env, 0, 0, 0, 0, 0, "", 0, 0, Name, EmptyNodeList, NavigatorID, NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                                Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                                Call cs8.GoNext()
                            Loop While cs8.OK
                        End If
                        Call cs8.Close()
                        '.
                        'CS = Main.OpenCSSQL("default", GetMenuSQL("((Parentid=0)or(Parentid is null))", NavigatorContentName))
                        'Do While Main.iscsok(CS)
                        '    Name = Trim(csx.getText("name"))
                        '    NavigatorID = csx.getInteger("ID")
                        '    NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name)
                        '    NodeIDString = CStr(NavigatorID)
                        '    If AutoManageAddons Then
                        '        '
                        '        ' special cases - root nodes that do not just deliver menu entries
                        '        '
                        '        Select Case LCase(Name)
                        '            Case "manage add-ons"
                        '                NodeIDString = NodeIDManageAddons
                        '            Case "settings"
                        '                NodeIDString = NodeIDSettings
                        '            Case "tools"
                        '                NodeIDString = NodeIDTools
                        '            Case "reports"
                        '                NodeIDString = NodeIDReports
                        '        End Select
                        '    End If
                        '    s = s & GetNavigatorNode(cp, 0, 0, 0, 0, 0, "", 0, 0, Name, EmptyNodeList, NavigatorID, NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, OpenNodeList, NodeIDString, NodeNavigatorJS,"")
                        '    Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        '    Call Main.NextCSRecord(CS)
                        'Loop
                        'Call Main.closeCs(CS)
                        '
                        ' Add a Legacy Menu node to the top-most parent menu at the very end
                        '
                        If cp.Utils.EncodeBoolean(cp.Site.GetText("AllowNavigatorLegacyEntry", "0")) Then
                            Name = "Legacy Menu"
                            NavIconTitleHtmlEncoded = "Legacy Menu"
                            NodeIDString = NodeIDLegacyMenu
                            s = s & GetNode(cp, env, 0, 0, 0, 0, 0, "", 0, 0, Name, EmptyNodeList, 0, NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, NodeTypeEnum.NodeTypeEntry, False, BlockSubNodes, OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                            Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        End If
                    Case NodeIDSettings
                        '
                        ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                        '
                        s = s & getNodeListMixed(cp, env, EmptyNodeList, TopParentNode, AutoManageAddons, NodeType, OpenNodeList, 3, cp.Content.GetRecordID("navigator entries", "settings"), NavIconTypeSetting, NodeNavigatorJS)
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    Case NodeIDTools
                        '
                        ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                        '
                        s = s & getNodeListMixed(cp, env, EmptyNodeList, TopParentNode, AutoManageAddons, NodeType, OpenNodeList, 4, cp.Content.GetRecordID("navigator entries", "tools"), NavIconTypeTool, NodeNavigatorJS)
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    Case NodeIDReports
                        '
                        ' list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                        '
                        s = s & getNodeListMixed(cp, env, EmptyNodeList, TopParentNode, AutoManageAddons, NodeType, OpenNodeList, 2, cp.Content.GetRecordID("navigator entries", "reports"), NavIconTypeReport, NodeNavigatorJS)
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                    Case Else
                        '
                        ' numeric node (default case) - list navigator records with parent=TopParentNode
                        '
                        Dim CS As Integer = -1
                        If IsNumeric(TopParentNode) Then
                            If InStr(1, EmptyNodeList & ",", "," & TopParentNode & ",") <> 0 Then
                                EmptyNodeList = EmptyNodeList
                            Else
                                '
                                ' Navigator Entries, child under TopParentNode
                                '
                                SQL = GetMenuSQL(cp, "parentid=" & TopParentNode, "")
                                BlockSubNodes = False
                                If Not csChildList.OpenSQL(SQL) Then
                                    '
                                    ' Empty list, add to EmptyNodeList
                                    '
                                    EmptyNodeList = EmptyNodeList & "," & TopParentNode
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
                        'Dim cs9 As CPCSBaseClass = cp.CSNew()
                        If csChildList.Open(NavigatorContentName, "id=" & TopParentNode) Then
                            ContentID = csChildList.GetInteger("ContentID")
                        End If
                        '
                        'CS = Main.openCSContent(NavigatorContentName, "id=" & TopParentNode)
                        'If Main.iscsok(CS) Then
                        '    ContentID = csx.getInteger("ContentID")
                        'End If
                        If ContentID <> 0 Then
                            ContentID = ContentID
                            Dim ContentName As String = cp.Content.GetRecordName("content", ContentID)
                            If ContentName <> "" Then
                                'ContentTableName =cp.Content.GetTable(ContentName)
                                csChildList.Close()
                                Dim Ptr As Integer = 0
                                If csChildList.Open(ContentName, "", "name", True, "ID,Name,ContentControlID", 20, 1) Then
                                    EmptyNodeList = Replace(EmptyNodeList, "," & TopParentNode, "")
                                    Do
                                        NavigatorID = csChildList.GetInteger("ID")
                                        Dim RecordName As String = csChildList.GetText("Name")
                                        If RecordName = "" Then
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
                                        s = s & cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & ATag & IconNoSubNodes & "</a>&nbsp;" & ATag & RecordName & "</a></div>"
                                        Ptr = Ptr + 1
                                        Call csChildList.GoNext()
                                    Loop While csChildList.OK() And Ptr < 20
                                    If Ptr = 20 Then
                                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Open All '" & NavigatorContentName & "'")
                                        Link = "?cid=" & ContentID
                                        Dim IconClosed As String
                                        s = s & cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconClosed & "&nbsp;<a href=""" & Link & """ title=""" & NavIconTitleHtmlEncoded & """>more...</a></div>"
                                    End If
                                End If
                                'CS = Main.openCSContent(ContentName, , , , , , "ID,Name,ContentControlID", 20, 1)
                                ' SQL = "select top 20 " _
                                '     & " ID,Name,ContentControlID from " & ContentTableName
                                ' CS = Main.OpenCSSQL("default", SQL)
                                'Ptr = 0
                                'If Main.iscsok(CS) Then
                                '    EmptyNodeList = Replace(EmptyNodeList, "," & TopParentNode, "")
                                '    Do While Main.iscsok(CS) And Ptr < 20
                                '        NavigatorID = csx.getInteger("ID")
                                '        RecordName = csx.getText("Name")
                                '        If RecordName = "" Then
                                '            RecordName = "Record " & NavigatorID
                                '        End If
                                '        '
                                '        If Len(RecordName) > 53 Then
                                '            RecordName = Left(RecordName, 25) & "..." & Right(RecordName, 25)
                                '        End If
                                '        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Edit '" & RecordName & "' in '" & ContentName & "'")
                                '        IconNoSubNodes = IconRecord
                                '        IconNoSubNodes = Replace(IconNoSubNodes, "{title}", NavIconTitleHtmlEncoded)
                                '        Link = "?id=" & NavigatorID & "&cid=" & csx.getInteger("ContentControlID") & "&af=4"
                                '        ATag = "<a href=""" & Link & """ title=""" & NavIconTitleHtmlEncoded & """>"
                                '        s = s & cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & ATag & IconNoSubNodes & "</a>&nbsp;" & ATag & RecordName & "</a></div>"
                                '        Ptr = Ptr + 1
                                '        Call Main.NextCSRecord(CS)
                                '    Loop
                                '    If Ptr = 20 Then
                                '        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Open All '" & NavigatorContentName & "'")
                                '        Link = "?cid=" & ContentID
                                '        s = s & cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconClosed & "&nbsp;<a href=""" & Link & """ title=""" & NavIconTitleHtmlEncoded & """>more...</a></div>"
                                '    End If
                                'End If
                            End If
                        End If
                    End If
                ElseIf csChildList.OK Then
                    '
                    ' List out child menus
                    '
                    Do
                        CollectionID = csChildList.GetInteger("CollectionID")
                        NavigatorID = csChildList.GetInteger("ID")
                        Name = Trim(csChildList.GetText("name"))
                        Dim NewWindow As Boolean = csChildList.GetBoolean("newwindow")
                        ContentID = csChildList.GetInteger("ContentID")
                        Link = Trim(csChildList.GetText("LinkPage"))
                        addonid = csChildList.GetInteger("AddonID")
                        NavIconType = csChildList.GetInteger("NavIconType")
                        NavIconTitle = csChildList.GetText("NavIconTitle")
                        Dim HelpAddonID As Integer = csChildList.GetInteger("HelpAddonID")
                        If HelpAddonID <> 0 Then
                            HelpAddonID = HelpAddonID
                        End If
                        Dim helpCollectionID As Integer = csChildList.GetInteger("HelpCollectionID")
                        If NavIconTitle = "" Then
                            NavIconTitle = Name
                        End If
                        ContentControlID = csChildList.GetInteger("ContentControlID")
                        If LCase(Name) = "all content" Then
                            '
                            ' special case: any Navigator Entry named 'all content' returns the content list
                            '
                            NodeIDString = NodeIDAllContentList
                        Else
                            NodeIDString = CStr(NavigatorID)
                        End If
                        linkSuffixList = ""
                        If (ContentID <> 0) And (env.isDeveloper) Then
                            linkSuffixList = "<a href=""" & env.contentFieldEditToolPrefix & ContentID & """>edit</a>"
                            If Not String.IsNullOrEmpty(linkSuffixList) Then
                                linkSuffixList = "&nbsp;(" & linkSuffixList & ")"
                            End If
                        End If
                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle)
                        Dim SettingPageID As Integer
                        s = s & GetNode(cp, env, CollectionID, ContentControlID, helpCollectionID, HelpAddonID, ContentID, Link, addonid, SettingPageID, Name, EmptyNodeList, NavigatorID, NavIconType, NavIconTitleHtmlEncoded, AutoManageAddons, NodeType, NewWindow, BlockSubNodes, OpenNodeList, NodeIDString, NodeNavigatorJS, linkSuffixList)
                        Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                        Call csChildList.GoNext()
                    Loop While csChildList.OK
                    Call csChildList.Close()
                End If
                '
                If EmptyNodeListInitial <> EmptyNodeList Then
                    Call cp.Cache.Save(cacheKey, EmptyNodeList, env.cacheDependencyList)
                End If
                returnNav = s
            Catch ex As Exception
                Call HandleError(cp, ex)
            End Try
            Return returnNav
        End Function
        '
        '========================================================================
        '   list mixed nodes (settings/reports/tools)
        '========================================================================
        '
        Friend Function getNodeListMixed(cp As CPBaseClass, env As NavigatorEnvironment, EmptyNodeList As String, TopParentNode As String, AutoManageAddons As Boolean, NodeType As NodeTypeEnum, OpenNodeList As String, AddonNavTypeID As Integer, MenuParentNodeID As Integer, AdminNavIconTypeSetting As Integer, ByRef Return_DraggableJS As String) As String
            Dim returnNav As String = ""
            Try
                Dim NodeDraggableJS As String
                Dim LastName As String
                Dim Index As New keyPtrIndexClass
                Dim SortNodes() As SortNodeType
                Dim SortPtr As Integer
                Dim testPtr As Integer
                Dim NodeIDString As String
                Dim SQL As String
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
                        If Not cp.User.IsDeveloper() Then
                            Criteria = Criteria & "and(admin<>0)"
                        End If
                    End If
                    Dim cs10 As CPCSBaseClass = cp.CSNew
                    If cs10.Open("add-ons", Criteria, "name") Then
                        Do
                            ReDim Preserve SortNodes(SortPtr)
                            With SortNodes(SortPtr)
                                .Name = Trim(cs10.GetText("name"))
                                .addonid = cs10.GetInteger("ID")
                                .NavIconTitle = .Name
                                .ContentControlID = cs10.GetInteger("ContentControlID")
                                .NavIconType = AdminNavIconTypeSetting
                                Call Index.SetPointer(.Name, SortPtr)
                            End With
                            SortPtr = SortPtr + 1
                            cs10.GoNext()
                        Loop While cs10.OK
                    End If
                    cs10.Close()
                    '
                    'CS = Main.openCSContent("Add-ons", Criteria, , , , , "*")
                    'SortPtr = 0
                    'Do While Main.iscsok(CS)
                    '    ReDim Preserve SortNodes(SortPtr)
                    '    With SortNodes(SortPtr)
                    '        .Name = Trim(csx.getText("name"))
                    '        .addonid = csx.getinteger( "ID")
                    '        .NavIconTitle = .Name
                    '        .ContentControlID = csx.getinteger( "ContentControlID")
                    '        .NavIconType = AdminNavIconTypeSetting
                    '        Call Index.SetPointer(.Name, SortPtr)
                    '    End With
                    '    SortPtr = SortPtr + 1
                    '    Call Main.NextCSRecord(CS)
                    'Loop
                    'Call Main.closeCs(CS)
                    '
                    ' Add real navigator nodes to node list
                    '
                    SQL = GetMenuSQL(cp, "parentid=" & MenuParentNodeID, "")
                    Dim cs11 As CPCSBaseClass = cp.CSNew
                    If cs11.OpenSQL(SQL) Then
                        Do
                            ReDim Preserve SortNodes(SortPtr)
                            With SortNodes(SortPtr)
                                .Name = Trim(cs11.GetText("name"))
                                .NavigatorID = cs11.GetInteger("ID")
                                .CollectionID = cs11.GetInteger("CollectionID")
                                .NewWindow = cs11.GetBoolean("newwindow")
                                .ContentID = cs11.GetInteger("ContentID")
                                .Link = Trim(cs11.GetText("LinkPage"))
                                .addonid = cs11.GetInteger("AddonID")
                                .NavIconType = cs11.GetInteger("NavIconType")
                                .NavIconTitle = cs11.GetText("NavIconTitle")
                                .HelpAddonID = cs11.GetInteger("HelpAddonID")
                                .helpCollectionID = cs11.GetInteger("HelpCollectionID")
                                If .NavIconTitle = "" Then
                                    .NavIconTitle = .Name
                                End If
                                .ContentControlID = cs11.GetInteger("ContentControlID")
                                If LCase(.Name) = "all content" Then
                                    '
                                    ' special case: any Navigator Entry named 'all content' returns the content list
                                    '
                                    .NodeIDString = NodeIDAllContentList
                                Else
                                    .NodeIDString = CStr(.NavigatorID)
                                    .NavIconTitle = .Name
                                    .ContentControlID = cs11.GetInteger("ContentControlID")
                                End If
                                Call Index.SetPointer(.Name, SortPtr)
                            End With
                            SortPtr = SortPtr + 1
                            Call cs11.GoNext()
                        Loop While cs11.OK
                    End If
                    Call cs11.Close()
                    '
                    ''
                    'CS = Main.OpenCSSQL("default", SQL)
                    'BlockSubNodes = False
                    'Do While Main.iscsok(CS)
                    '    ReDim Preserve SortNodes(SortPtr)
                    '    With SortNodes(SortPtr)
                    '        .Name = Trim(csx.getText("name"))
                    '        .NavigatorID = csx.getinteger("ID")
                    '        .CollectionID = csx.getinteger("CollectionID")
                    '        .NewWindow = csx.getBoolean("newwindow")
                    '        .ContentID = csx.getinteger("ContentID")
                    '        .Link = Trim(csx.getText("LinkPage"))
                    '        .addonid = csx.getinteger("AddonID")
                    '        .NavIconType = csx.getinteger("NavIconType")
                    '        .NavIconTitle = csx.getText("NavIconTitle")
                    '        .HelpAddonID = csx.getinteger("HelpAddonID")
                    '        .helpCollectionID = csx.getinteger("HelpCollectionID")
                    '        If .NavIconTitle = "" Then
                    '            .NavIconTitle = .Name
                    '        End If
                    '        .ContentControlID = csx.getinteger("ContentControlID")
                    '        If LCase(.Name) = "all content" Then
                    '            '
                    '            ' special case: any Navigator Entry named 'all content' returns the content list
                    '            '
                    '            .NodeIDString = NodeIDAllContentList
                    '        Else
                    '            .NodeIDString = CStr(.NavigatorID)
                    '            .NavIconTitle = .Name
                    '            .ContentControlID = csx.getinteger("ContentControlID")
                    '        End If
                    '        Call Index.SetPointer(.Name, SortPtr)
                    '    End With
                    '    SortPtr = SortPtr + 1
                    '    Call Main.NextCSRecord(CS)
                    'Loop
                    'Call Main.closeCs(CS)
                    testPtr = Index.GetFirstPointer()
                    If testPtr = -1 Then
                        '
                        ' Empty list, add to EmptyNodeList
                        '
                        EmptyNodeList = EmptyNodeList & "," & TopParentNode
                    Else
                        LastName = ""
                        Do While testPtr <> -1
                            SortPtr = cp.Utils.EncodeInteger(testPtr)
                            With SortNodes(SortPtr)
                                If LCase(.Name) <> LastName Then
                                    returnNav = returnNav & GetNode(cp, env, .CollectionID, .ContentControlID, .helpCollectionID, .HelpAddonID, .ContentID, .Link, .addonid, 0, .Name, EmptyNodeList, .NavigatorID, .NavIconType, cp.Utils.EncodeHTML(.NavIconTitle), AutoManageAddons, NodeType, .NewWindow, BlockSubNodes, OpenNodeList, .NodeIDString, NodeDraggableJS, "")
                                    Return_DraggableJS = Return_DraggableJS & NodeDraggableJS
                                    LastName = LCase(.Name)
                                End If
                            End With
                            testPtr = Index.GetNextPointer()
                        Loop
                    End If
                End If
            Catch ex As Exception
                HandleError(cp, ex)
            End Try
            Return returnNav
        End Function
        '
        '
        '
        Private Function GetNode(cp As CPBaseClass, env As NavigatorEnvironment, CollectionID As Integer, ContentControlID As Integer, helpCollectionID As Integer, HelpAddonID As Integer, ContentID As Integer, Link As String, addonid As Integer, ignore As Integer, Name As String, EmptyNodeList As String, NavigatorID As Integer, NavIconType As Integer, NavIconTitleHtmlEncoded As String, AutoManageAddons As Boolean, NodeType As NodeTypeEnum, NewWindow As Boolean, BlockSubNodes As Boolean, OpenNodeList As String, NodeIDString As String, ByRef Return_NavigatorJS As String, linkSuffixList As String) As String
            On Error GoTo ErrorTrap
            '
            Dim csCollection As Integer
            Dim collectionName As String
            Dim collectionHelpLink As String
            Dim collectionHelp As String
            Dim NodeNavigatorJS As String
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
            IsVisible = (CollectionID <> 0) Or (helpCollectionID <> 0) Or (HelpAddonID <> 0) Or (ContentID <> 0) Or (Link <> "") Or (addonid <> 0) Or (LCase(WorkingName) = "all content") Or (LCase(WorkingName) = "add-ons with no collection")
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
                    If addonid <> 0 Then
                        '
                        ' link to addon
                        '
                        Link = ""
                        Dim cs12 As CPCSBaseClass = cp.CSNew()
                        If cs12.Open("Add-ons", "id=" & addonid, "", True, "remotemethod,name,ccguid") Then
                            AddonGuid = cs12.GetText("ccguid")
                            AddonName = cs12.GetText("name")
                            If cs12.GetBoolean("remotemethod") Then
                                NewWindow = True
                                Link = cp.Site.GetText("adminUrl") & "?" & RequestNameRemoteMethodAddon & "=" & cp.Utils.EncodeRequestVariable(AddonName)
                            End If

                        End If
                        Call cs12.Close()
                        ''
                        'CSAddon = Main.OpenCSContentRecord("Add-ons", addonid, , , "remotemethod,name,ccguid")
                        'If Main.iscsok(CSAddon) Then
                        '    AddonGuid = Main.getcsText(CSAddon, "ccguid")
                        '    AddonName = Main.getcsText(CSAddon, "name")
                        '    If Main.GetCSBoolean(CSAddon, "remotemethod") Then
                        '        NewWindow = True
                        '        Link = cp.Site.GetText("adminUrl") & "?" & RequestNameRemoteMethodAddon & "=" & cp.Utils.EncodeRequestVariable(AddonName)
                        '    End If
                        'End If
                        'Call Main.closeCs(CSAddon)
                        If Link = "" Then
                            If AddonGuid <> "" Then
                                Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "addonguid", AddonGuid, True)
                            Else
                                Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "addonid", CStr(addonid), True)
                            End If
                        End If
                        NavLinkHTMLId = "a" & addonid
                        workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                        If NewWindow Then
                            workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & Link & """ target=""_blank"" title=""Run '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
                        Else
                            workingNameHtmlEncoded = "<a class=navDrag name=navLink id=""" & NavLinkHTMLId & """ href=""" & Link & """ title=""Run '" & workingNameHtmlEncoded & "'"">" & workingNameHtmlEncoded & "</a>"
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
                        Dim cs13 As CPCSBaseClass = cp.CSNew
                        If Not cs13.Open("add-on collections", "id=" & helpCollectionID, "name", True, "name,helpLink,help") Then
                            BlockNode = True
                        Else
                            collectionName = cs13.GetText("name")
                            collectionHelpLink = cs13.GetText("helpLink")
                            collectionHelp = cs13.GetText("help")
                            '
                            WorkingName = collectionName
                            Link = collectionHelpLink
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                            If Link <> "" Then
                                NewWindow = True
                            ElseIf (collectionHelp <> "") Then
                                Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "helpcollectionid", CStr(helpCollectionID), True)
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
                        cs13.Close()
                        '
                        'csCollection = Main.openCSContent("add-on collections", "id=" & helpCollectionID, , , , , "name,helpLink,help")
                        'If Not Main.iscsok(csCollection) Then
                        '    BlockNode = True
                        'Else
                        '    collectionName = cs13.getText( "name")
                        '    collectionHelpLink = cs13.getText( "helpLink")
                        '    collectionHelp = cs13.getText( "help")
                        '    '
                        '    WorkingName = collectionName
                        '    Link = collectionHelpLink
                        '    workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                        '    If Link <> "" Then
                        '        NewWindow = True
                        '    ElseIf (collectionHelp <> "") Then
                        '        Link = cp.Utils.ModifyLinkQueryString(cp.Site.GetText("adminUrl"), "helpcollectionid", CStr(helpCollectionID), True)
                        '    Else
                        '        BlockNode = True
                        '    End If
                        '    If Not BlockNode Then
                        '        If NewWindow Then
                        '            workingNameHtmlEncoded = "<a href=""" & Link & """ target=""_blank"" title=""Help for Collection '" & workingNameHtmlEncoded & "'"">Help</a>"
                        '        Else
                        '            workingNameHtmlEncoded = "<a href=""" & Link & """ title=""Help for Collection '" & workingNameHtmlEncoded & "'"">Help</a>"
                        '        End If
                        '    End If
                        'End If
                        'Call Main.closeCs(csCollection)
                    Else
                        workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName)
                    End If
                End If
                '
                If Not BlockNode Then
                    If BlockSubNodes Or (addonid <> 0) Then
                        '
                        ' This is a hardcoded item (like Add-on), it has no subnodes
                        '
                        s = s & cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
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
                            s = s & cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                        ElseIf InStr(1, OpenNodeList & ",", "," & NodeIDString & ",") <> 0 Then
                            '
                            ' This node is open
                            '
                            SubNav = GetNodeList(cp, env, NodeIDString, OpenNodeList, NodeNavigatorJS)
                            Return_NavigatorJS = Return_NavigatorJS & NodeNavigatorJS
                            If SubNav <> "" Then
                                '
                                ' display the subnav
                                '
                                s = s & cr & "<div class=ccNavLink ID=" & DivIDClosed & " style=""display:none;""><A class=""ccNavClosed"" href=""#"" onclick=""AdminNavOpenClick('" & DivIDClosed & "','" & DivIDOpened & "','" & DivIDContent & "','" & NodeIDString & "','" & DivIDEmpty & "');return false;"">" & IconClosed & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                s = s & cr & "<div class=ccNavLink ID=" & DivIDOpened & "><A class=""ccNavOpened"" href=""#"" onclick=""AdminNavCloseClick('" & DivIDOpened & "','" & DivIDClosed & "','" & DivIDContent & "','" & NodeIDString & "');return false;"">" & IconOpened & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                s = s _
                                    & cr & "<div class=ccNavLinkChild ID=" & DivIDContent & ">" _
                                    & kmaIndent(SubNav) _
                                    & cr & "</div>"
                            Else
                                '
                                ' it has a NO subnav
                                '
                                s = s & cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            End If
                        Else
                            '
                            ' This node is closed
                            '
                            s = s & cr & "<div class=ccNavLink ID=" & DivIDClosed & " ><A class=""ccNavClosed"" href=""#"" onclick=""AdminNavOpenClick('" & DivIDClosed & "','" & DivIDOpened & "','" & DivIDContent & "','" & NodeIDString & "','','" & DivIDContent & "');return false;"">" & IconClosed & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            s = s & cr & "<div class=ccNavLink ID=" & DivIDOpened & " style=""display:none;""><A class=""ccNavOpened"" href=""#"" onclick=""AdminNavCloseClick('" & DivIDOpened & "','" & DivIDClosed & "','" & DivIDContent & "','" & NodeIDString & "');return false;"">" & IconOpened & "</A>&nbsp;" & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            s = s & cr & "<div class=""ccNavLink ccNavLinkEmpty"" ID=" & DivIDEmpty & " style=""display:none;"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                            s = s & cr & "<div class=ccNavLinkChild ID=" & DivIDContent & " style=""display:none;margin-left:20px;"">&nbsp;&nbsp;&nbsp;&nbsp;<img src=""/cclib/images/ajax-loader-small.gif"" width=""16"" height=""16""></div>"
                        End If
                    End If
                    If NavLinkHTMLId <> "" Then
                        'Return_NavigatorJS = Return_NavigatorJS _
                        '    & "jQuery('#" & NavLinkHTMLId & "').draggable({" _
                        '        & "opacity: 0.50" _
                        '        & ",helper: 'clone'" _
                        '        & ",revert: 'invalid'" _
                        '        & ",stop: function(event, ui){" _
                        '            & "navDrop('" & NavLinkHTMLId & "',ui.position.left,ui.position.top);" _
                        '        & "}" _
                        '        & ",cursor: 'move'" _
                        '    & "});"
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
        '========================================================================
        ' Get sql for menu
        '   if MenuContentName is blank, it will select values from either cdef
        '========================================================================
        '
        Private Function GetMenuSQL(cp As CPBaseClass, ParentCriteria As String, MenuContentName As String) As String
            On Error GoTo ErrorTrap
            '
            Dim CMCriteria As String
            'Dim ParentCriteria As String
            Dim Criteria As String
            Dim SQL As String
            Dim ContentControlCriteria As String
            Dim SelectList As String
            Dim ContentManagementList As String
            '
            Criteria = "(Active<>0)"
            If MenuContentName <> "" Then
                Criteria = Criteria & "AND" & cp.Content.GetContentControlCriteria(MenuContentName)
            End If
            'ParentCriteria = KmaEncodeMissingText(ParentCriteria, "")
            If cp.User.IsDeveloper() Then
                '
                ' ----- Developer
                '
            ElseIf cp.User.IsAdmin Then
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
                ContentManagementList = GetContentManagementList(cp)
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
            If ParentCriteria <> "" Then
                Criteria = "(" & ParentCriteria & ")AND" & Criteria
            End If
            SelectList = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, ccMenuEntries.AddonID, ccMenuEntries.NavIconType, ccMenuEntries.NavIconTitle, HelpAddonID,HelpCollectionID,0 as collectionid"
            GetMenuSQL = "select " & SelectList & " from ccMenuEntries where " & Criteria & " order by ccMenuEntries.Name"
            Call cp.Site.TestPoint("adminNavigator, getmenuSql=" & GetMenuSQL)
            Exit Function
            '
            ' ----- Error Trap
            '
ErrorTrap:
            Call cp.Site.ErrorReport("Trap")
            '
        End Function
        ''
        ''
        ''
        'Private Sub HandleClassAppendLogfile(MethodName As String, Context As String)
        '    Call AppendLogFile2(Main.ApplicationName, Context, "ccWeb3", "AddonManClass", MethodName, 0, "", "", False, True, Main.ServerLink, "", "trace")

        'End Sub
        ''
        ''===========================================================================
        ''
        ''===========================================================================
        ''
        'Private Sub cp.Site.ErrorReport((MethodName As String, Optional Context As String = "")
        '    '
        '    If Main Is Nothing Then
        '        Call HandleError2("unknown", Context, App.EXEName, "AddonManClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, "unknown")
        '    Else
        '        Call HandleError2(Main.ApplicationName, Context, App.EXEName, "AddonManClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, Main.ServerLink)
        '    End If
        '    '
        'End Sub
        ''
        '===========================================================================
        '   Get Authoring List
        '       returns a comma delimited list of ContentIDs that the Member can author
        '===========================================================================
        '
        Private Function GetContentManagementList(cp As CPBaseClass) As String
            On Error GoTo ErrorTrap
            '
            Dim SQL As String
            'Dim RS As Recordset
            Dim CIDArray() As Object
            Dim CIDCount As Integer
            Dim CIDPointer As Integer
            'Dim cdef As CDefType
            Dim ContentName As String
            Dim ContentID As Integer
            Dim ChildIDList As String
            Dim sqlTablememberRules As String
            Dim cs As CPCSBaseClass = cp.CSNew
            '
            sqlTablememberRules = cp.Content.GetTable("Member Rules")
            '
            SQL = "Select ccGroupRules.ContentID as ID" _
                & " FROM ((" & sqlTablememberRules _
                & " Left Join ccGroupRules on " & sqlTablememberRules & ".GroupID=ccGroupRules.GroupID)" _
                & " Left Join ccContent on ccGroupRules.ContentID=ccContent.ID)" _
                & " WHERE" _
                    & " (" & sqlTablememberRules & ".MemberID=" & cp.User.Id & ")" _
                    & " AND(ccGroupRules.Active<>0)" _
                    & " AND(ccContent.Active<>0)" _
                    & " AND(" & sqlTablememberRules & ".Active<>0)"
            If cs.OpenSQL(SQL) Then
                Do
                    ContentID = cs.GetInteger("id")
                    GetContentManagementList = GetContentManagementList & "," & CStr(ContentID)
                    ContentName = cp.Content.GetRecordName("content", ContentID)
                    If ContentName <> "" Then
                        ChildIDList = cp.Content.GetProperty(ContentName, "ChildIDList")
                        If ChildIDList <> "" Then
                            GetContentManagementList = GetContentManagementList & "," & ChildIDList
                        End If
                    End If
                    Call cs.GoNext()
                Loop While cs.OK
            End If
            Call cs.Close()
            ''
            'RS = Main.openRSSQL("Default", SQL)
            'If IsRSOK(RS) Then
            '    CIDArray = RS.GetRows()
            '    CIDCount = UBound(CIDArray, 2) + 1
            'End If
            'RS = Nothing
            'For CIDPointer = 0 To CIDCount - 1
            '    ContentID = cp.Utils.EncodeInteger(CIDArray(0, CIDPointer))
            '    GetContentManagementList = GetContentManagementList & "," & CStr(ContentID)
            '    ContentName = Main.GetContentNameByID(ContentID)
            '    If ContentName <> "" Then
            '        ChildIDList = Main.GetContentProperty(ContentName, "ChildIDList")
            '        'cdef = Main.GetContentDefinition(ContentName)
            '        If ChildIDList <> "" Then
            '            GetContentManagementList = GetContentManagementList & "," & ChildIDList
            '        End If
            '    End If
            'Next
            If Len(GetContentManagementList) > 1 Then
                GetContentManagementList = Mid(GetContentManagementList, 2)
            End If
            '
            Exit Function
ErrorTrap:
            Call cp.Site.ErrorReport("")
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
        '
        '========================================================================
        '   HandleError
        '       Logs the error and either resumes next, or raises it to the next level
        '========================================================================
        '
        Public Sub HandleError(cp As CPBaseClass, ByVal ex As Exception, ByVal className As String, ByVal methodName As String, ByVal cause As String)
            Try
                Dim errMsg As String = className & "." & methodName & ", cause=[" & cause & "], ex=[" & ex.ToString & "]"
                cp.Site.ErrorReport(ex, errMsg)
            Catch exIgnore As Exception
                '
            End Try
        End Sub
        '
        Public Sub HandleError(cp As CPBaseClass, ByVal ex As Exception)
            Dim frame As StackFrame = New StackFrame(1)
            Dim method As System.Reflection.MethodBase = frame.GetMethod()
            Dim type As System.Type = method.DeclaringType()
            Dim methodName As String = method.Name
            '
            Call HandleError(cp, ex, type.Name, methodName, "unexpected exception")
        End Sub
    End Class
End Namespace

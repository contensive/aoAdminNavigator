
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Contensive.AdminNavigator
    Public Class NodeListMixedController
        '
        Private Shared Function compare1(a As KeyValuePair(Of String, SortNodeType), b As KeyValuePair(Of String, SortNodeType)) As Integer
            Return a.Key.CompareTo(b.Key)
        End Function
        '
        '====================================================================================================
        '========================================================================
        ' list mixed nodes (settings/reports/tools)   
        Friend Shared Function getNodeListMixed(cp As CPBaseClass, env As NavigatorEnvironment, xEmptyNodeList As String, TopParentNode As String, AutoManageAddons As Boolean, NodeType As NodeTypeEnum, xOpenNodeList As String, AddonNavTypeID As Integer, MenuParentNodeID As Integer, AdminNavIconTypeSetting As Integer, ByRef Return_DraggableJS As String) As String
            Dim returnNav As String = ""
            Try
                Dim NodeDraggableJS As String
                Dim lastName As String
                'Dim Index As New keyPtrIndexClass
                'Dim SortNodes() As SortNodeType
                Dim sortedNodes As New List(Of KeyValuePair(Of String, SortNodeType))
                Dim SortPtr As Integer
                Dim testPtr As Integer
                Dim NodeIDString As String
                Dim SQL As String
                Dim Criteria As String
                Dim BlockSubNodes As Boolean
                '
                ' list mixed nodes (settings/reports/tools), includes menu nodes and addons with type='setting' sorted in
                '
                If InStr(1, env.EmptyNodeList & ",", "," & TopParentNode & ",") <> 0 Then
                    env.EmptyNodeList = env.EmptyNodeList
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
                            Dim nodeName As String = Trim(cs10.GetText("name"))
                            Dim sortNode As New SortNodeType() With {
                                    .Name = Trim(cs10.GetText("name")),
                                    .addonid = cs10.GetInteger("ID"),
                                    .NavIconTitle = .Name,
                                    .ContentControlID = cs10.GetInteger("ContentControlID"),
                                    .NavIconType = AdminNavIconTypeSetting
                                }
                            sortedNodes.Add(New KeyValuePair(Of String, SortNodeType)(nodeName, sortNode))
                            SortPtr += 1
                            cs10.GoNext()
                        Loop While cs10.OK
                    End If
                    cs10.Close()
                    '
                    ' Add real navigator nodes to node list
                    '
                    Using cs11 As CPCSBaseClass = cp.CSNew
                        If cs11.OpenSQL(MenuSqlController.GetMenuSQL(cp, "parentid=" & MenuParentNodeID, "")) Then
                            Do
                                Dim nodeName As String = Trim(cs11.GetText("name"))
                                Dim navIconTitle As String = cs11.GetText("NavIconTitle")
                                Dim sortNode As New SortNodeType() With {
                                    .Name = Trim(cs11.GetText("name")),
                                    .NavigatorID = cs11.GetInteger("ID"),
                                    .CollectionID = cs11.GetInteger("CollectionID"),
                                    .NewWindow = cs11.GetBoolean("newwindow"),
                                    .contentID = cs11.GetInteger("ContentID"),
                                    .Link = Trim(cs11.GetText("LinkPage")),
                                    .addonid = cs11.GetInteger("AddonID"),
                                    .NavIconType = cs11.GetInteger("NavIconType"),
                                    .NavIconTitle = If(String.IsNullOrEmpty(navIconTitle), nodeName, navIconTitle),
                                    .HelpAddonID = cs11.GetInteger("HelpAddonID"),
                                    .helpCollectionID = cs11.GetInteger("HelpCollectionID"),
                                    .ContentControlID = cs11.GetInteger("ContentControlID"),
                                    .NodeIDString = If(nodeName.ToLowerInvariant() = "all content", NodeIDAllContentList, NodeIDAllContentList)
                                }
                                If (nodeName.ToLowerInvariant() = "all content") Then
                                    sortNode.NodeIDString = NodeIDAllContentList
                                Else
                                    sortNode.NodeIDString = sortNode.NavigatorID.ToString()
                                    sortNode.NavIconTitle = nodeName
                                End If
                                sortedNodes.Add(New KeyValuePair(Of String, SortNodeType)(nodeName, sortNode))
                                SortPtr = SortPtr + 1
                                Call cs11.GoNext()
                            Loop While cs11.OK
                        End If
                        Call cs11.Close()
                    End Using
                    '
                    'sortedNodes = sortedNodes.Sort(compare1)
                    sortedNodes.Sort(Function(valueA, valueB) valueA.Key.CompareTo(valueB.Key))
                    '
                    If sortedNodes.Count = 0 Then
                        '
                        ' Empty list, add to env.EmptyNodeList
                        '
                        env.EmptyNodeList = env.EmptyNodeList & "," & TopParentNode
                    Else
                        lastName = ""
                        Dim contentNodeList As String = ""
                        For Each kvp In sortedNodes
                            Dim nodeName As String = kvp.Key.ToLowerInvariant()
                            If nodeName <> lastName Then
                                Dim sortedNode As SortNodeType = kvp.Value
                                If (Not sortedNode.contentID.Equals(0)) Then
                                    '
                                    ' -- content nodes last
                                    contentNodeList &= NodeController.GetNode(cp, env, sortedNode.CollectionID, sortedNode.ContentControlID, sortedNode.helpCollectionID, sortedNode.HelpAddonID, sortedNode.contentID, sortedNode.Link, Nothing, 0, sortedNode.Name, env.EmptyNodeList, sortedNode.NavigatorID, sortedNode.NavIconType, cp.Utils.EncodeHTML(sortedNode.NavIconTitle), AutoManageAddons, NodeType, sortedNode.NewWindow, BlockSubNodes, env.OpenNodeList, sortedNode.NodeIDString, NodeDraggableJS, "")
                                ElseIf (sortedNode.addonid.Equals(0)) Then
                                    '
                                    ' -- hard-coded nodes
                                    returnNav &= NodeController.GetNode(cp, env, sortedNode.CollectionID, sortedNode.ContentControlID, sortedNode.helpCollectionID, sortedNode.HelpAddonID, sortedNode.contentID, sortedNode.Link, Nothing, 0, sortedNode.Name, env.EmptyNodeList, sortedNode.NavigatorID, sortedNode.NavIconType, cp.Utils.EncodeHTML(sortedNode.NavIconTitle), AutoManageAddons, NodeType, sortedNode.NewWindow, BlockSubNodes, env.OpenNodeList, sortedNode.NodeIDString, NodeDraggableJS, "")
                                Else
                                    '
                                    ' -- addon nodes next
                                    Dim addon As AddonModel = DbBaseModel.create(Of AddonModel)(cp, sortedNode.addonid)
                                    returnNav &= NodeController.GetNode(cp, env, sortedNode.CollectionID, sortedNode.ContentControlID, sortedNode.helpCollectionID, sortedNode.HelpAddonID, sortedNode.contentID, sortedNode.Link, addon, 0, sortedNode.Name, env.EmptyNodeList, sortedNode.NavigatorID, sortedNode.NavIconType, cp.Utils.EncodeHTML(sortedNode.NavIconTitle), AutoManageAddons, NodeType, sortedNode.NewWindow, BlockSubNodes, env.OpenNodeList, sortedNode.NodeIDString, NodeDraggableJS, "")
                                End If
                                Return_DraggableJS &= NodeDraggableJS
                                lastName = nodeName
                            End If
                        Next
                        returnNav &= contentNodeList
                    End If
                End If
            Catch ex As Exception

                cp.Site.ErrorReport(ex)
            End Try
            Return returnNav
        End Function

        'Private Shared Function compare1(x As KeyValuePair(Of String, SortNodeType), y As KeyValuePair(Of String, SortNodeType)) As Integer
        '    Throw New NotImplementedException()
        'End Function
        '
        '
    End Class
End Namespace



Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    Public NotInheritable Class NodeManageAddonsController
        '
        '====================================================================================================
        '
        Public Shared Function getNode(cp As CPBaseClass, env As ApplicationEnvironmentModel, ByRef Return_NavigatorJS As String, ByRef nodeIDString As String) As String
            Try
                nodeIDString = ""
                Dim result As String = ""
                Dim addonManagerVerionGuid As String = "{4DD876E7-BCC4-4AF0-B32E-59FFAB816478}"
                Dim addon As Models.Db.AddonModel = Models.Db.DbBaseModel.create(Of Models.Db.AddonModel)(cp, addonManagerVerionGuid)
                If (addon IsNot Nothing) Then
                    result &= NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "?addonguid=" & addonManagerVerionGuid, addon, 0, "Add-on Manager", env.emptyNodeList, 0, NavIconTypeAddon, "Add-on Manager", env.autoManageAddons, NodeTypeEnum.NodeTypeAddon, False, False, env.openNodeList, nodeIDString, Return_NavigatorJS, "")
                End If
                '
                ' List Collections
                Dim fieldList As String = "Name,0 as id,ccaddoncollections.id as collectionid,0 as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," & NavIconTypeFolder & " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid,blockNavigatorNode,system"
                Dim criteria As String = "((system=0)or(system is null))"
                '
                ' -- reduce confusion
                criteria &= "and((blockNavigatorNode=0)or(blockNavigatorNode is null))"
                'If Not env.isDeveloper Then
                '    criteria &= "and((blockNavigatorNode=0)or(blockNavigatorNode is null))"
                'End If
                Using cs3 As CPCSBaseClass = cp.CSNew()
                    Dim NodeType As NodeTypeEnum = NodeTypeEnum.NodeTypeCollection
                    Dim BlockSubNodes As Boolean = False
                    If cs3.Open("Add-on Collections", criteria, "name", True, fieldList, 9999, 1) Then
                        Do
                            Dim Name As String = Trim(cs3.GetText("name"))
                            Dim NavIconTitle As String = Name
                            Dim CollectionID As Integer = cs3.GetInteger("collectionid")
                            nodeIDString = NodeIDManageAddonsCollectionPrefix & "." & CollectionID
                            Dim NavIconTitleHtmlEncoded As String = cp.Utils.EncodeHTML(NavIconTitle)
                            Dim linkSuffixList As String = ""
                            '
                            ' -- reduce confusion
                            'If env.isDeveloper Then
                            '    linkSuffixList &= "<a href=""" & env.addonEditCollectionUrlPrefix & CollectionID & """>edit</a>"
                            '    If cs3.GetBoolean("system") Then
                            '        linkSuffixList &= ",sys"
                            '    End If
                            '    If cs3.GetBoolean("blockNavigatorNode") Then
                            '        linkSuffixList &= ",dev"
                            '    End If
                            '    linkSuffixList = "&nbsp;(" & linkSuffixList & ")"
                            'End If
                            result &= NodeController.getNode(cp, env, CollectionID, 0, 0, 0, 0, "", Nothing, 0, Name, env.emptyNodeList, 0, NavIconTypeAddon, NavIconTitleHtmlEncoded, env.autoManageAddons, NodeTypeEnum.NodeTypeCollection, False, False, env.openNodeList, nodeIDString, Return_NavigatorJS, linkSuffixList)
                            Return_NavigatorJS &= Return_NavigatorJS
                            Call cs3.GoNext()
                        Loop While cs3.OK()
                    End If
                    Call cs3.Close()
                End Using
                '
                ' Advanced folder to contain edit links to create addons and collections
                '
                nodeIDString = NodeIDManageAddonsAdvanced
                result &= NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "", Nothing, 0, "Advanced", env.emptyNodeList, 0, NavIconTypeFolder, "Add-ons With No Collection", env.autoManageAddons, NodeTypeEnum.NodeTypeEntry, False, False, env.openNodeList, nodeIDString, Return_NavigatorJS, "")
                Return_NavigatorJS &= Return_NavigatorJS
                Return result
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
    End Class
End Namespace


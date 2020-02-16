
Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    '
    Public Class NavigatorEnvironment
        Public adminUrl As String
        Public isDeveloper As Boolean
        Public buildVersion As String
        Public addonEditAddonUrlPrefix As String
        Public addonEditCollectionUrlPrefix As String
        Public contentFieldEditToolPrefix As String
        Public cacheDependencyList As String
        Public allowToolTip As Boolean
        ''' <summary>
        ''' 
        ''' </summary>
        Public EmptyNodeList As String = ""
        '
        Public OpenNodeList As String = ""
        '
        Public AutoManageAddons As Boolean = False
        '
        Public AdminNavOpen As Boolean
        '
        Public Sub New(cp As CPBaseClass)
            adminUrl = cp.Site.GetText("adminurl")
            allowToolTip = cp.Site.GetBoolean("Admin Navigator Allow Tooltip")
            buildVersion = cp.Site.GetText("buildversion")
            isDeveloper = cp.User.IsDeveloper()
            If isDeveloper Then
                addonEditAddonUrlPrefix = adminUrl & "?cid=" & cp.Content.GetID("add-ons") & "&af=4&id="
                addonEditCollectionUrlPrefix = adminUrl & "?cid=" & cp.Content.GetID("add-on collections") & "&af=4&id="
                contentFieldEditToolPrefix = adminUrl & "?af=105&contentid="
            End If
            OpenNodeList = cp.Visit.GetText("AdminNavOpenNodeList", "")
            AdminNavOpen = cp.Utils.EncodeBoolean(cp.Visit.GetText("AdminNavOpen", "1"))


            cacheDependencyList = C51CacheController.createDependencyKeyInvalidateOnChange(cp, "ccMenuEntries", "default") & "," & C51CacheController.createDependencyKeyInvalidateOnChange(cp, "ccaggregatefunctions", "default")
        End Sub
    End Class

End Namespace

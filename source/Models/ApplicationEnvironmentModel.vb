
Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    '
    Public Class ApplicationEnvironmentModel
        Implements IDisposable
        '
        Private cp As CPBaseClass
        '
        ''' <summary>
        ''' normalized admin url, leading slash, no trailing slash
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property adminUrl As String
            Get
                If adminUrl_local IsNot Nothing Then Return adminUrl_local
                adminUrl_local = cp.Site.GetText("adminurl")
                If String.IsNullOrEmpty(adminUrl_local) Then Return ""
                If adminUrl_local.Substring(0, 1) <> "/" Then adminUrl_local = "/" & adminUrl_local
                If adminUrl_local.Length = 1 Then Return adminUrl_local
                Do While adminUrl_local.Length > 0 AndAlso adminUrl_local.Substring(adminUrl_local.Length - 1, 1) = "/"
                    adminUrl_local = adminUrl_local.Substring(0, adminUrl_local.Length - 1)
                Loop
                Return adminUrl_local
            End Get
        End Property
        Private adminUrl_local As String = Nothing

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
            Me.cp = cp
            '
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
            cacheDependencyList = cp.Cache.CreateDependencyKeyInvalidateOnChange("ccMenuEntries", "default") & "," & cp.Cache.CreateDependencyKeyInvalidateOnChange("ccaggregatefunctions", "default")
        End Sub
        '
#Region " IDisposable Support "
        Protected disposed As Boolean = False
        '
        '==========================================================================================
        ''' <summary>
        ''' dispose
        ''' </summary>
        ''' <param name="disposing"></param>
        ''' <remarks></remarks>
        Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposed Then
                If disposing Then
                    '
                    ' ----- call .dispose for managed objects
                    '
                End If
                '
                ' Add code here to release the unmanaged resource.
                '
            End If
            Me.disposed = True
        End Sub
        ' Do not change or add Overridable to these methods.
        ' Put cleanup code in Dispose(ByVal disposing As Boolean).
        Public Overloads Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overrides Sub Finalize()
            Dispose(False)
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace

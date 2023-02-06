
Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    '
    Public Class ApplicationEnvironmentModel
        Implements IDisposable
        '
        Private ReadOnly cp As CPBaseClass
        '
        '==========================================================================================
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
        '
        '==========================================================================================
        ''' <summary>
        ''' The user. this was added because contensive returns cp.user.developer true for admin (fixed 11/11/2021)
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property user As Models.Db.PersonModel
            Get
                If user_local IsNot Nothing Then Return user_local
                user_local = Models.Db.DbBaseModel.create(Of Models.Db.PersonModel)(cp, cp.User.Id)
                Return user_local
            End Get
        End Property
        Private user_local As Models.Db.PersonModel = Nothing
        '
        '==========================================================================================
        ''' <summary>
        ''' is the current user a administrator
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property isAdministrator As Boolean
            Get
                Return user.admin Or user.developer
            End Get
        End Property
        '
        '==========================================================================================
        ''' <summary>
        ''' is the current user a developer
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property isDeveloper As Boolean
            Get
                Return user.developer
            End Get
        End Property
        '
        '==========================================================================================
        Public ReadOnly Property buildVersion As String
            Get
                If buildVersion_local IsNot Nothing Then Return buildVersion_local
                buildVersion_local = cp.Site.GetText("buildversion")
                Return buildVersion_local
            End Get
        End Property
        Private buildVersion_local As String = Nothing
        '
        '==========================================================================================
        Public ReadOnly Property addonEditAddonUrlPrefix As String
            Get
                If addonEditAddonUrlPrefix_local IsNot Nothing Then Return addonEditAddonUrlPrefix_local
                addonEditAddonUrlPrefix_local = adminUrl & "?cid=" & cp.Content.GetID("add-ons") & "&af=4&id="
                Return addonEditAddonUrlPrefix_local
            End Get
        End Property
        Private addonEditAddonUrlPrefix_local As String = Nothing
        '
        '==========================================================================================
        Public ReadOnly Property addonEditCollectionUrlPrefix As String
            Get
                If addonEditCollectionUrlPrefix_local IsNot Nothing Then Return addonEditCollectionUrlPrefix_local
                addonEditCollectionUrlPrefix_local = adminUrl & "?cid=" & cp.Content.GetID("add-on collections") & "&af=4&id="
                Return addonEditCollectionUrlPrefix_local
            End Get
        End Property
        Private addonEditCollectionUrlPrefix_local As String = Nothing
        '
        '==========================================================================================
        Public ReadOnly Property contentFieldEditToolPrefix As String
            Get
                If contentFieldEditToolPrefix_local IsNot Nothing Then Return contentFieldEditToolPrefix_local
                contentFieldEditToolPrefix_local = adminUrl & "?af=105&contentid="
                Return contentFieldEditToolPrefix_local
            End Get
        End Property
        Private contentFieldEditToolPrefix_local As String = Nothing
        '
        '==========================================================================================
        Public ReadOnly Property cacheDependencyList As String
            Get
                If cacheDependencyList_local IsNot Nothing Then Return cacheDependencyList_local
                addonEditCollectionUrlPrefix_local = ""
                cacheDependencyList_local = cp.Cache.CreateTableDependencyKey("ccMenuEntries", "default") & "," & cp.Cache.CreateTableDependencyKey("ccaggregatefunctions", "default")
                Return cacheDependencyList_local
            End Get
        End Property
        Private cacheDependencyList_local As String = Nothing
        '
        '==========================================================================================
        Public ReadOnly Property allowToolTip As Boolean
            Get
                Return cp.Site.GetBoolean("Admin Navigator Allow Tooltip")
            End Get
        End Property
        '
        Private ReadOnly Property emptyNodeListCacheKey As String
            Get
                If isDeveloper Then Return "AdminNav EmptyNodeList Dev"
                If isAdministrator Then Return "AdminNav EmptyNodeList Admin"
                Return "AdminNav EmptyNodeList CM" & cp.User.Id
            End Get
        End Property
        '
        '==========================================================================================
        ''' <summary>
        ''' 
        ''' </summary>
        Public Property emptyNodeList As String
            Get
                If emptyNodeList_local IsNot Nothing Then Return emptyNodeList_local
                '
                ' -- read cache version, return if not empty
                emptyNodeList_local = cp.Cache.GetText(emptyNodeListCacheKey)
                If Not String.IsNullOrEmpty(emptyNodeList_local) Then Return emptyNodeList_local
                '
                ' -- build a new string, store in cache and return
                Dim SQL As String = "select n.ID from ccMenuEntries n left join ccMenuEntries c on c.parentid=n.id Where c.ID Is Null group by n.id"
                Dim cs2 As CPCSBaseClass = cp.CSNew()
                If cs2.OpenSQL(SQL) Then
                    Do
                        emptyNodeList_local &= "," & cs2.GetText("id")
                        Call cs2.GoNext()
                    Loop While cs2.OK()
                    emptyNodeList_local = emptyNodeList_local.Substring(1)
                End If
                Call cs2.Close()
                Call cp.Cache.Store(emptyNodeListCacheKey, emptyNodeList_local, cacheDependencyList)
                Return emptyNodeList_local
            End Get
            Set(value As String)
                Call cp.Cache.Store(emptyNodeListCacheKey, value, cacheDependencyList)
            End Set
        End Property
        Private emptyNodeList_local As String = Nothing
        '
        '==========================================================================================
        '
        Public Property openNodeList As String
            Get
                If openNodeList_local IsNot Nothing Then Return openNodeList_local
                openNodeList_local = cp.Visit.GetText("AdminNavOpenNodeList", "")
                Return openNodeList_local
            End Get
            Set(value As String)
                Call cp.Visit.SetProperty("AdminNavOpenNodeList", value)
            End Set
        End Property
        Private openNodeList_local As String = Nothing
        '
        '==========================================================================================
        '
        Public Property autoManageAddons As Boolean = False
        '
        '==========================================================================================
        '
        Public ReadOnly Property adminNavOpen As Boolean
            Get
                Return cp.Utils.EncodeBoolean(cp.Visit.GetText("AdminNavOpen", "1"))
            End Get
        End Property
        '
        Public ReadOnly Property allowAdminNavSaveState As Boolean
            Get
                Return cp.Site.GetBoolean("adminNav allowSaveState", True)
            End Get
        End Property
        '
        '==========================================================================================
        '
        Public Sub New(cp As CPBaseClass)
            Me.cp = cp
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

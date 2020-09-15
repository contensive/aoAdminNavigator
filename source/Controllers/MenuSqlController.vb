
Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    Public NotInheritable Class MenuSqlController
        '
        '====================================================================================================
        ''' <summary>
        ''' Get sql for menu, if MenuContentName is blank, it will select values from either cdef
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="ParentCriteria"></param>
        ''' <param name="MenuContentName"></param>
        ''' <returns></returns>
        Public Shared Function getMenuSQL(cp As CPBaseClass, ParentCriteria As String, MenuContentName As String) As String
            Try
                '
                Dim Criteria As String = "(Active<>0)"
                If MenuContentName <> "" Then
                    Criteria = Criteria & "AND" & cp.Content.GetContentControlCriteria(MenuContentName)
                End If
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
                    Dim ContentManagementList As String = MenuSqlController.GetContentManagementList(cp)
                    '
                    Dim CMCriteria As String
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
                Dim SelectList As String = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, ccMenuEntries.AddonID, ccMenuEntries.NavIconType, ccMenuEntries.NavIconTitle, HelpAddonID,HelpCollectionID,0 as collectionid"
                getMenuSQL = "select " & SelectList & " from ccMenuEntries where " & Criteria & " order by ccMenuEntries.Name"
                Call cp.Site.TestPoint("adminNavigator, getmenuSql=" & getMenuSQL)
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
        '
        '===========================================================================
        ''' <summary>
        ''' Get Authoring List, returns a comma delimited list of ContentIDs that the Member can author
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <returns></returns>
        Private Shared Function GetContentManagementList(cp As CPBaseClass) As String
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
    End Class
End Namespace


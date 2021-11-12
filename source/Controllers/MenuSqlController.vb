
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
        ''' <param name="menuContentName"></param>
        ''' <returns></returns>
        Public Shared Function getMenuSQL(cp As CPBaseClass, ParentCriteria As String, menuContentName As String) As String
            Try
                Dim criteria As String = "(Active<>0)"
                If Not String.IsNullOrEmpty(menuContentName) Then
                    criteria &= "AND" & cp.Content.GetContentControlCriteria(menuContentName)
                End If
                If cp.User.IsDeveloper() Then
                    '
                    ' -- Developer
                ElseIf cp.User.IsAdmin Then
                    '
                    ' -- Administrator
                    criteria = criteria _
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
                    Dim ContentManagementList As String = MenuSqlController.getContentManagementList(cp)
                    '
                    Dim CMCriteria As String
                    If String.IsNullOrEmpty(ContentManagementList) Then
                        CMCriteria = "(1=0)"
                    Else
                        'ContentManagementList = Mid(ContentManagementList, 2, Len(ContentManagementList) - 2)
                        If InStr(1, ContentManagementList, ",") = 0 Then
                            CMCriteria = "(ccContent.ID=" & ContentManagementList & ")"
                        Else
                            CMCriteria = "(ccContent.ID in (" & ContentManagementList & "))"
                        End If
                    End If

                    criteria = criteria _
                    & "AND((DeveloperOnly is null)or(DeveloperOnly=0))" _
                    & "AND((AdminOnly is null)or(AdminOnly=0))" _
                    & "AND(ID in (" _
                    & " SELECT AllowedEntries.ID" _
                    & " FROM CCMenuEntries AllowedEntries LEFT JOIN ccContent ON AllowedEntries.ContentID = ccContent.ID" _
                    & " Where (" & CMCriteria & "and(ccContent.Active<>0)And((ccContent.DeveloperOnly is null)or(ccContent.DeveloperOnly=0))And((ccContent.AdminOnly is null)or(ccContent.AdminOnly=0)))" _
                        & "OR(ccContent.ID Is Null)" _
                    & "))"
                End If
                If Not String.IsNullOrEmpty(ParentCriteria) Then
                    criteria = "(" & ParentCriteria & ")AND" & criteria
                End If
                Dim SelectList As String = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, ccMenuEntries.AddonID, ccMenuEntries.NavIconType, ccMenuEntries.NavIconTitle, HelpAddonID,HelpCollectionID,0 as collectionid"
                getMenuSQL = "select " & SelectList & " from ccMenuEntries where " & criteria & " order by ccMenuEntries.Name"
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
        Private Shared Function getContentManagementList(cp As CPBaseClass) As String
            Try
                Dim result As String = ""
                Dim sqlTablememberRules As String = cp.Content.GetTable("Member Rules")
                Dim SQL As String = "Select ccGroupRules.ContentID as ID" _
                    & " FROM ((" & sqlTablememberRules _
                    & " Left Join ccGroupRules on " & sqlTablememberRules & ".GroupID=ccGroupRules.GroupID)" _
                    & " Left Join ccContent on ccGroupRules.ContentID=ccContent.ID)" _
                    & " WHERE" _
                    & " (" & sqlTablememberRules & ".MemberID=" & cp.User.Id & ")" _
                    & " AND(ccGroupRules.Active<>0)" _
                    & " AND(ccContent.Active<>0)" _
                    & " AND(" & sqlTablememberRules & ".Active<>0)"
                Using cs As CPCSBaseClass = cp.CSNew
                    If cs.OpenSQL(SQL) Then
                        Do
                            Dim ContentID As Integer = cs.GetInteger("id")
                            result &= "," & CStr(ContentID)
                            Dim contentName As String = cp.Content.GetRecordName("content", ContentID)
                            If Not String.IsNullOrEmpty(contentName) Then
                                Dim ChildIDList As String = cp.Content.GetProperty(contentName, "ChildIDList")
                                If Not String.IsNullOrEmpty(ChildIDList) Then
                                    result &= "," & ChildIDList
                                End If
                            End If
                            Call cs.GoNext()
                        Loop While cs.OK
                    End If
                    Call cs.Close()
                End Using
                If Len(result) > 1 Then result = Mid(result, 2)
                '
                Return result
            Catch ex As Exception
                Call cp.Site.ErrorReport("")
                Throw
            End Try
        End Function
    End Class
End Namespace


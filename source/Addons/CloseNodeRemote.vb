
Imports System.Linq
Imports Contensive.BaseClasses

Namespace Contensive.AdminNavigator
    Public Class CloseNodeRemote
        Inherits AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Try
                Dim nodeId As String = CP.Doc.GetText("nodeid")
                If (Not String.IsNullOrWhiteSpace(nodeId)) Then
                    Dim nodeList As List(Of String) = CP.Visit.GetText("AdminNavOpenNodeList", "").Split(","c).ToList()
                    If (nodeList.Contains(nodeId)) Then
                        nodeList.Remove(nodeId)
                        Call CP.Visit.SetProperty("AdminNavOpenNodeList", String.Join(",", nodeList))
                    End If
                End If
                Return String.Empty
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
    End Class
End Namespace

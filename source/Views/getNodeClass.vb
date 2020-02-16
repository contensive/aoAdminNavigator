
Option Explicit On
Option Strict On

Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Contensive.AdminNavigator
    Public Class GetNodeClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        ''' <summary>
        ''' Return the full navigator
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Return getNode(CP, New NavigatorEnvironment(CP))
        End Function
        '
        '====================================================================================================
        '
        Public Shared Function getNode(ByVal CP As CPBaseClass, env As NavigatorEnvironment) As String
            Try
                Dim parentNode As String = CP.Doc.GetText("nodeid")
                If (parentNode <> "") Then
                    If String.IsNullOrEmpty(env.OpenNodeList) Then
                        env.OpenNodeList = "," & parentNode
                        Call CP.Visit.SetProperty("AdminNavOpenNodeList", env.OpenNodeList)
                    Else
                        If ((env.OpenNodeList & ",").IndexOf("," & parentNode.ToString & ",") < 0) Then
                            env.OpenNodeList = env.OpenNodeList & "," & parentNode
                            Call CP.Visit.SetProperty("AdminNavOpenNodeList", env.OpenNodeList)
                        End If
                    End If
                End If
                Dim returnJs As String = ""
                Dim returnHtml As String = NodeListController.getNodeList(CP, env, parentNode, returnJs)
                If Not String.IsNullOrEmpty(returnJs) Then
                    returnJs = "if(window.navDrop) {" & returnJs & "};"
                    Call CP.Doc.AddBodyEnd("<script language=""javascript"" type=""text/javascript/"">jQuery(document).ready(function(){" & returnJs & "})</script>")
                End If
                Return returnHtml
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Return String.Empty
            End Try
        End Function
    End Class
End Namespace

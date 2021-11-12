
Imports Contensive.BaseClasses

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
            Return getNode(CP, New ApplicationEnvironmentModel(CP))
        End Function
        '
        '====================================================================================================
        '
        Public Shared Function getNode(ByVal CP As CPBaseClass, env As ApplicationEnvironmentModel) As String
            Try
                Dim parentNode As String = CP.Doc.GetText("nodeid")
                If Not String.IsNullOrEmpty(parentNode) Then
                    If String.IsNullOrEmpty(env.openNodeList) Then
                        env.openNodeList = "," & parentNode
                    Else
                        If ((env.openNodeList & ",").IndexOf("," & parentNode.ToString & ",") < 0) Then
                            env.openNodeList = env.openNodeList & "," & parentNode
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
                Throw
            End Try
        End Function
    End Class
End Namespace

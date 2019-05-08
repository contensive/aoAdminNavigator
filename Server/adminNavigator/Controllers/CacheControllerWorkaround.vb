
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions
Imports Contensive.BaseClasses

Namespace Contensive.adminNavigator
    '
    ' -- simulate C51 base class
    Public Class C51CacheController
        '
        '====================================================================================================
        '
        Public Shared Sub updateLastModified(cp As CPBaseClass, tableName As String, dataSourceName As String)
            cp.Cache.Save(createDependencyKeyInvalidateOnChange(cp, tableName, dataSourceName), Now.ToString())
        End Sub
        '
        '====================================================================================================
        '
        Public Shared Function createDependencyKeyInvalidateOnChange(cp As CPBaseClass, tableName As String, dataSourceName As String) As String
            Dim key As String = "lastrecordmodifieddate/" & If(String.IsNullOrWhiteSpace(dataSourceName), "default/" & tableName.Trim().ToLowerInvariant() & "/", dataSourceName.Trim().ToLowerInvariant() & "/" & tableName.Trim().ToLowerInvariant() & "/")
            Return Regex.Replace(key, "0x[a-fA-F\\d]{2}", "_").ToLowerInvariant().Replace(" ", "_")
        End Function
        '
    End Class
End Namespace


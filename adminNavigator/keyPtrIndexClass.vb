
'Imports Contensive.Core.ccCommonModule
'Imports Contensive.Core.cpCommonUtilsClass
Imports System.Runtime.InteropServices

'
Namespace Contensive.adminNavigator
    '
    '
    '
    '<ComVisible(True)> _
    '<Guid("6C9D6ED7-BEEA-323C-8379-2B4D3827B0AD")> _
    Public Class keyPtrIndexClass
        '
        ' new serializable and deserialize
        '   declare a private instance of a class that holds everything
        '   keyPtrIndex uses the class
        '   call serialize on keyPtrIndex to json serialize the storage object and return the string
        '   call deserialise to populate the storage object from the argument
        '
        ' ----- Index Type - This structure is the basis for Element Indexing
        '       Records are read into thier data structure, and keys(Key,ID,etc.) and pointers
        '       are put in the KeyPointerArrays.
        '           AddIndex( Key, value )
        '           BubbleSort( Index ) - sorts the index by the key field
        '           GetIndexValue( index, Key ) - retrieves the pointer
        '
        ' These  GUIDs provide the COM identity for this class 
        ' and its COM interfaces. If you change them, existing 
        ' clients will no longer be able to access the class.
        'Public Const ClassId As String = "BB8AFA32-1C0A-4CDB-BE3B-D9E6AA91A656"
        'Public Const InterfaceId As String = "353333D8-FB3B-4340-B8B6-C5547B46F5DF"
        'Public Const EventsId As String = "1407C7AD-08DF-44DB-898E-7B3CB9F86EB3"
        '
        Private Const KeyPointerArrayChunk = 1000
        '
        Public Class storageClass
            '
            Public ArraySize As Integer
            Public ArrayCount As Integer
            Public ArrayDirty As Boolean
            Public UcaseKeyArray() As String
            Public PointerArray() As String
            Public ArrayPointer As Integer
        End Class
        '
        Private store As New storageClass
        '
        '
        '
        Public Function ExportPropertyBag() As String
            Dim returnBag As String = ""
            Try
                Dim json As New System.Web.Script.Serialization.JavaScriptSerializer
                '

                returnBag = json.Serialize(store)
                'returnBag = Newtonsoft.Json.JsonConvert.SerializeObject(store)
                'Catch ex As Newtonsoft.Json.JsonException
                '    Throw New indexException("ExportPropertyBag JSON error", ex)
            Catch ex As Exception
                Throw New indexException("ExportPropertyBag error", ex)
            End Try
            Return returnBag
        End Function
        '
        '
        '
        Public Sub ImportPropertyBag(ByVal bag As String)
            Try
                Dim json As New System.Web.Script.Serialization.JavaScriptSerializer
                '
                store = json.Deserialize(Of storageClass)(bag)
                'store = Newtonsoft.Json.JsonConvert.DeserializeObject(Of storageClass)(bag)
                'Catch ex As Newtonsoft.Json.JsonException
                '    Throw New indexException("ImportPropertyBag JSON error", ex)
            Catch ex As Exception
                Throw New indexException("ImportPropertyBag error", ex)
            End Try
        End Sub
        '
        '========================================================================
        '   Returns a pointer into the index for this Key
        '   Used only by GetIndexValue and setIndexValue
        '   Returns -1 if there is no match
        '========================================================================
        '
        Private Function GetArrayPointer(ByVal Key As String) As Integer
            Dim ArrayPointer As Integer = -1
            Try
                Dim UcaseTargetKey As String
                'Dim ElementKey As String
                Dim HighGuess As Integer
                Dim LowGuess As Integer
                Dim PointerGuess As Integer
                Dim test As String
                test = ""
                '
                If store.ArrayDirty Then
                    Call Sort()
                End If
                '
                ArrayPointer = -1
                If store.ArrayCount > 0 Then
                    UcaseTargetKey = Replace(UCase(Key), vbCrLf, "")
                    LowGuess = -1
                    HighGuess = store.ArrayCount - 1
                    Do While (HighGuess - LowGuess) > 1
                        PointerGuess = CInt(Int((HighGuess + LowGuess) / 2))
                        If UcaseTargetKey = store.UcaseKeyArray(PointerGuess) Then
                            HighGuess = PointerGuess
                            Exit Do
                        ElseIf UcaseTargetKey < store.UcaseKeyArray(PointerGuess) Then
                            HighGuess = PointerGuess
                        Else
                            LowGuess = PointerGuess
                        End If
                    Loop
                    If UcaseTargetKey = store.UcaseKeyArray(HighGuess) Then
                        ArrayPointer = HighGuess
                    End If
                End If

            Catch ex As Exception
                Throw New indexException("getArrayPointer error", ex)
            End Try
            Return ArrayPointer
        End Function
        '
        '========================================================================
        '   Returns the matching pointer from a ContentIndex
        '   Returns -1 if there is no match
        '========================================================================
        '
        Public Function GetPointer(ByVal Key As String) As Integer
            Dim returnKey As Integer = -1
            Try
                Dim test As String
                Dim MatchFound As Boolean
                Dim UcaseKey As String
                test = ""
                '
                UcaseKey = Replace(UCase(Key), vbCrLf, "")
                'UcaseKey = UCase(Key)
                store.ArrayPointer = GetArrayPointer(Key)
                If store.ArrayPointer > -1 Then
                    ' Make sure this is the first match
                    MatchFound = True
                    Do While MatchFound
                        store.ArrayPointer = store.ArrayPointer - 1
                        If store.ArrayPointer < 0 Then
                            MatchFound = False
                        Else
                            MatchFound = (store.UcaseKeyArray(store.ArrayPointer) = UcaseKey)
                        End If
                    Loop
                    store.ArrayPointer = store.ArrayPointer + 1
                    returnKey = kmaEncodeInteger(store.PointerArray(store.ArrayPointer))
                End If
            Catch ex As Exception
                Throw New indexException("GetPointer error", ex)
            End Try
            Return returnKey
        End Function
        '
        '========================================================================
        '   Add an element to an ContentIndex
        '
        '   if the entry is a duplicate, it is added anyway
        '========================================================================
        '
        Public Sub SetPointer(ByVal Key As String, ByVal Pointer As Integer)
            Try
                Dim keyToSave As String
                '
                keyToSave = Replace(UCase(Key), vbCrLf, "")
                '
                If store.ArrayCount >= store.ArraySize Then
                    store.ArraySize = store.ArraySize + KeyPointerArrayChunk
                    ReDim Preserve store.PointerArray(store.ArraySize)
                    ReDim Preserve store.UcaseKeyArray(store.ArraySize)
                End If
                store.ArrayPointer = store.ArrayCount
                store.ArrayCount = store.ArrayCount + 1
                store.UcaseKeyArray(store.ArrayPointer) = keyToSave
                store.PointerArray(store.ArrayPointer) = CStr(Pointer)
                store.ArrayDirty = True
            Catch ex As Exception
                Throw New indexException("SetPointer error", ex)
            End Try
        End Sub
        '
        '========================================================================
        '   Returns the next matching pointer from a ContentIndex
        '   Returns -1 if there is no match
        '========================================================================
        '
        Public Function GetNextPointerMatch(ByVal Key As String) As Integer
            Dim nextPointerMatch As Integer = -1
            Try
                Dim UcaseKey As String
                '
                If store.ArrayPointer < (store.ArrayCount - 1) Then
                    store.ArrayPointer = store.ArrayPointer + 1
                    UcaseKey = UCase(Key)
                    If (store.UcaseKeyArray(store.ArrayPointer) = UcaseKey) Then
                        nextPointerMatch = kmaEncodeInteger(store.PointerArray(store.ArrayPointer))
                    Else
                        store.ArrayPointer = store.ArrayPointer - 1
                    End If
                End If
            Catch ex As Exception
                Throw New indexException("GetNextPointerMatch error", ex)
            End Try
            Return nextPointerMatch
        End Function
        '
        '========================================================================
        '   Returns the first Pointer in the current index
        '   returns empty if there are no Pointers indexed
        '========================================================================
        '
        Public Function GetFirstPointer() As Integer
            Dim firstPointer As Integer = -1
            Try
                If store.ArrayDirty Then
                    Call Sort()
                End If
                '
                ' GetFirstPointer = -1
                If store.ArrayCount > 0 Then
                    store.ArrayPointer = 0
                    firstPointer = kmaEncodeInteger(store.PointerArray(store.ArrayPointer))
                End If
                '
            Catch ex As Exception
                Throw New indexException("GetFirstPointer error", ex)
            End Try
            Return firstPointer
        End Function
        '
        '========================================================================
        '   Returns the next Pointer, past the last one returned
        '   Returns empty if the index is at the end
        '========================================================================
        '
        Public Function GetNextPointer() As Integer
            Dim nextPointer As Integer = -1
            Try
                If store.ArrayDirty Then
                    Call Sort()
                End If
                '
                'nextPointer = -1
                If (store.ArrayPointer + 1) < store.ArrayCount Then
                    store.ArrayPointer = store.ArrayPointer + 1
                    nextPointer = kmaEncodeInteger(store.PointerArray(store.ArrayPointer))
                End If
            Catch ex As Exception
                Throw New indexException("GetPointer error", ex)
            End Try
            Return nextPointer
        End Function
        '
        '========================================================================
        '
        '========================================================================
        '
        Private Sub BubbleSort()
            Try
                Dim TempUcaseKey As String
                Dim tempPtrString As String
                'Dim TempPointer as integer
                Dim CleanPass As Boolean
                Dim MaxPointer As Integer
                Dim SlowPointer As Integer
                Dim FastPointer As Integer
                Dim test As String
                Dim PointerDelta As Integer
                test = ""
                '
                If store.ArrayCount > 1 Then
                    PointerDelta = 1
                    MaxPointer = store.ArrayCount - 2
                    For SlowPointer = MaxPointer To 0 Step -1
                        CleanPass = True
                        For FastPointer = MaxPointer To (MaxPointer - SlowPointer) Step -1
                            If store.UcaseKeyArray(FastPointer) > store.UcaseKeyArray(FastPointer + PointerDelta) Then
                                TempUcaseKey = store.UcaseKeyArray(FastPointer + PointerDelta)
                                tempPtrString = store.PointerArray(FastPointer + PointerDelta)
                                store.UcaseKeyArray(FastPointer + PointerDelta) = store.UcaseKeyArray(FastPointer)
                                store.PointerArray(FastPointer + PointerDelta) = store.PointerArray(FastPointer)
                                store.UcaseKeyArray(FastPointer) = TempUcaseKey
                                store.PointerArray(FastPointer) = tempPtrString
                                CleanPass = False
                            End If
                        Next
                        If CleanPass Then
                            Exit For
                        End If
                    Next
                End If
                store.ArrayDirty = False
            Catch ex As Exception
                Throw New indexException("BubbleSort error", ex)
            End Try
        End Sub
        '
        '========================================================================
        '
        ' Made by Michael Ciurescu (CVMichael from vbforums.com)
        ' Original thread: http://www.vbforums.com/showthread.php?t=231925
        '
        '========================================================================
        '
        Private Sub QuickSort()
            Try
                If store.ArrayCount >= 2 Then
                    Call QuickSort_Segment(store.UcaseKeyArray, store.PointerArray, 0, store.ArrayCount - 1)
                End If
            Catch ex As Exception
                Throw New indexException("QuickSort error", ex)
            End Try
        End Sub
        '
        '
        '========================================================================
        '
        ' Made by Michael Ciurescu (CVMichael from vbforums.com)
        ' Original thread: http://www.vbforums.com/showthread.php?t=231925
        '
        '========================================================================
        '
        Private Sub QuickSort_Segment(ByVal C() As String, ByVal P() As String, ByVal First As Integer, ByVal Last As Integer)
            Try
                Dim Low As Integer, High As Integer
                Dim MidValue As String
                Dim TC As String
                Dim TP As String
                '
                Low = First
                High = Last
                MidValue = C((First + Last) \ 2)
                '
                Do
                    While C(Low) < MidValue
                        Low = Low + 1
                    End While
                    While C(High) > MidValue
                        High = High - 1
                    End While
                    If Low <= High Then
                        TC = C(Low)
                        TP = P(Low)
                        C(Low) = C(High)
                        P(Low) = P(High)
                        C(High) = TC
                        P(High) = TP
                        Low = Low + 1
                        High = High - 1
                    End If
                Loop While Low <= High
                If First < High Then
                    QuickSort_Segment(C, P, First, High)
                End If
                If Low < Last Then
                    QuickSort_Segment(C, P, Low, Last)
                End If
            Catch ex As Exception
                Throw New indexException("QuickSort_Segment error", ex)
            End Try
        End Sub
        '
        '
        '
        Private Sub Sort()
            Try
                Call QuickSort()
                store.ArrayDirty = False
            Catch ex As Exception
                Throw New indexException("Sort error", ex)
            End Try
        End Sub
    End Class
    '
    '
    '

    Public Class indexException
        Inherits System.Exception
        Implements System.Runtime.Serialization.ISerializable
        '
        'Private _message As String

        '


        Public Sub New()
            MyBase.New()
            ' Add implementation.
        End Sub

        Public Sub New(ByVal message As String)
            MyBase.New(message)
            ' Add implementation.
        End Sub

        Public Sub New(ByVal message As String, ByVal inner As Exception)
            MyBase.New(message, inner)
            ' Add implementation.
        End Sub

        ' This constructor is needed for serialization.
        Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
            MyBase.New(info, context)
            ' Add implementation.
        End Sub
    End Class

End Namespace
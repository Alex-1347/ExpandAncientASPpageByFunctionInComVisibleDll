Imports System.Text.RegularExpressions
<Runtime.InteropServices.ComVisible(True)>
<Runtime.InteropServices.Guid("24209092-7b45-446a-9b18-f8ec75f894e7")>
<Runtime.InteropServices.ClassInterface(Runtime.InteropServices.ClassInterfaceType.AutoDispatch)>
<Runtime.InteropServices.ProgId("ClassicAspService.InsertAdvTag")>
Public Class AdvBlockInsert
    Implements IAdvBlockInsert, IDisposable

    Private disposedValue As Boolean

    Public Function InsertTag(Txt As String) As String Implements IAdvBlockInsert.InsertTag
        Try
            Dim Lines As String() = IO.File.ReadAllLines(IO.Path.Combine(My.Application.Info.DirectoryPath, "Blocks.txt"))
            Dim InsertBlocks As New Dictionary(Of Integer, String)
            For Each OneLine As String In Lines
                Dim Line1 As String() = OneLine.Split(",")
                InsertBlocks.Add(CInt(Line1(0)) - 1, Line1(1))
            Next
            Dim PSearch As Regex = New Regex("</p.*?", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
            Dim Ret1 As New Text.StringBuilder(Txt)
            Dim Tmp1 As New Text.StringBuilder()
            Dim Ptags As MatchCollection
            For Each OneKey As Integer In InsertBlocks.Keys
                Ptags = PSearch.Matches(Ret1.ToString)
                If OneKey >= Ptags.Count Then Exit For
                Debug.Print($"Search key:{OneKey},value:{InsertBlocks(OneKey)},Ptags.Count:{Ptags.Count}")
                If InsertBlocks.ContainsKey(OneKey) Then
                    Debug.Print($"Insert:{OneKey}/{Ret1.Length}")
                    Tmp1.Append(Left(Ret1.ToString, Ptags(OneKey).Index + 4))
                    Tmp1.Append($"<!--{OneKey + 1}-->")
                    Tmp1.Append(InsertBlocks.Item(OneKey))
                    Tmp1.Append(Mid(Ret1.ToString, Ptags(OneKey).Index + 5))
                    Ret1 = New Text.StringBuilder(Tmp1.ToString)
                    Tmp1 = New Text.StringBuilder()
                    Continue For
                End If
            Next
            Tmp1 = Nothing : PSearch = Nothing : InsertBlocks = Nothing : Lines = Nothing : Ptags = Nothing
            Return Ret1.ToString
        Catch ex As Exception
            Return ex.Message & vbCrLf & Txt
        End Try
    End Function

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects)
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
            ' TODO: set large fields to null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

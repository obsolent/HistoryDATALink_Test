Imports System.Reflection

Public Class Form1
    Dim MyInterfaceName = "iImageFileInformation"
    Dim HistoryData As iImageFileInformation
    Dim x As DialogResult


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim FilePath As String = "C:\Users\obsol_000\Documents\~~VB_Projects\HistoryDATALink\HistoryDATALink\bin\Debug\HistoryDataLink.dll"
        Dim oAssembly As System.Reflection.Assembly
        oAssembly = [Assembly].LoadFrom(FilePath)

        ''Gets all referenced Types of an assembly that implement a specific interface
        Dim otherAssemblyTypes As IEnumerable(Of Type) =
          oAssembly.GetTypes().Where(Function(y) IsImplementing(y, MyInterfaceName))

        Call ChooseFile()
        Stop
    End Sub

    Private Function IsImplementing(objType As Type, strInterface As String) As Boolean

        Dim objInterface As Type

        If objType.IsPublic = False Then Return False

        'Ignore abstract classes
        If ((objType.Attributes And TypeAttributes.Abstract) = TypeAttributes.Abstract) Then Return False

        'See if this type implements our interface
        objInterface = objType.GetInterface(strInterface, True)
        If objInterface Is Nothing Then Return False

        'It does
        'Plugin = New AvailablePlugin
        '    Plugin.AssemblyPath = objDLL.Location
        '    Plugin.ClassName = objType.FullName
        '    Plugins.Add(Plugin)
        Return True

    End Function
    Private Sub ChooseFile()

        Do
            Dim FD = New OpenFileDialog
            With FD
                .InitialDirectory = "C:\Users\obsol_000\Pictures"
                .Multiselect = False
                x = .ShowDialog
                If x = DialogResult.OK Then

                    Dim lblChosenFile = .FileName
                    Dim lblBestMatch = ""
                    Dim lblEventPeriod = ""
                    Dim lblDescription = ""

                    Dim MatchOK As Boolean = HistoryData.GetHistoryData(.FileName, HistoryData.ImageFileInfo)
                    If MatchOK Then
                        With HistoryData.ImageFileInfo
                            lblBestMatch = .MatchedFile
                            lblEventPeriod = .EventPeriod
                            lblDescription = .Description
                        End With
                    End If
                Else
                    If x = DialogResult.Cancel Then End
                End If
                .Dispose()
            End With
        Loop

        GC.Collect()
    End Sub


End Class


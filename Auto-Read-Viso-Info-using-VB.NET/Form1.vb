Imports System.IO
Imports Microsoft.Office.Interop.Visio

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Define path of CSV file
        Dim csvFile As String = "Path\To\CSV"
        Dim outFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(csvFile, False)

        Dim docs As New Collection
        Dim topFolder As New DirectoryInfo("Path\To\vsd\Folder")

        outFile.WriteLine("Title, Name")

        For Each currentFile In topFolder.EnumerateFiles("*.vsd", SearchOption.AllDirectories)
            docs.Add(currentFile.FullName)
        Next
        For Each d In docs
            Dim visio As New Microsoft.Office.Interop.Visio.InvisibleApp
            Dim doc As Document
            doc = visio.Documents.Add(d)
            Dim filename = System.IO.Path.GetFileNameWithoutExtension(d)
            Dim p As Page
            For Each p In doc.Pages
                Dim n As String
                n = filename & "::" & p.Name
                outFile.WriteLine(p.Name & " ," & p.Name)
            Next
            doc.Close()
            visio.Quit()
        Next
        outFile.Close()
    End Sub
End Class

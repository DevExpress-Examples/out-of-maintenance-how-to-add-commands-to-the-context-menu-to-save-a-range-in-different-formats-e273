Imports Microsoft.VisualBasic
#Region "#usingsinuriprovider"
Imports System
Imports System.IO
Imports DevExpress.Office.Services
Imports DevExpress.Office.Utils
Imports DevExpress.XtraRichEdit.Utils
#End Region ' #usingsinuriprovider
Namespace GetTextMethods
#Region "#uriprovider"
    Public Class MyUriProvider
        Implements IUriProvider
        Private rootDirectory As String
        Public Sub New(ByVal rootDirectory As String)
            If String.IsNullOrEmpty(rootDirectory) Then
                Exceptions.ThrowArgumentException("rootDirectory", rootDirectory)
            End If
            Me.rootDirectory = rootDirectory
        End Sub

        Public Function CreateCssUri(ByVal rootUri As String, ByVal styleText As String, ByVal relativeUri As String) As String _
            Implements IUriProvider.CreateCssUri
            ' This method is not called when HTML content is obtained via the Document.GetHtmlText method.
            ' Styles are palced within the <style> tag. 
            Return String.Empty
        End Function

        Public Function CreateImageUri(ByVal rootUri As String, ByVal image As DevExpress.Office.Utils.OfficeImage, ByVal relativeUri As String) As String _
            Implements IUriProvider.CreateImageUri
            Dim imagesDir As String = String.Format("{0}\{1}", Me.rootDirectory, rootUri.Trim("/"c))
            If (Not Directory.Exists(imagesDir)) Then
                Directory.CreateDirectory(imagesDir)
            End If
            Dim imageName As String = String.Format("{0}\{1}.png", imagesDir, Guid.NewGuid())
            Dim bytes As Byte() = image.GetImageBytesSafe(OfficeImageFormat.Png)
            Using stream As FileStream = New FileStream(imageName, FileMode.Create, FileAccess.Write)
                stream.Write(bytes, 0, bytes.Length)
                stream.Close()
            End Using
            Return GetRelativePath(imageName)
        End Function

        Private Function GetRelativePath(ByVal path As String) As String
            Dim substring As String = path.Substring(Me.rootDirectory.Length)
            Return substring.Replace("\", "/").Trim("/"c)
        End Function
    End Class
#End Region ' #uriprovider
End Namespace

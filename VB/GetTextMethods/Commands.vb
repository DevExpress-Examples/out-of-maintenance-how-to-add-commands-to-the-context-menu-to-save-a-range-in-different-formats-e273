Imports Microsoft.VisualBasic
#Region "#usingsincommands"
Imports System
Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Commands
Imports DevExpress.XtraRichEdit.Export.Html
Imports System.Windows.Media
#End Region ' #usingsincommands
Namespace GetTextMethods
    #Region "#basecommand"
    Friend MustInherit Class MyFileSaveAsCommand
        Inherits SaveDocumentAsCommand
        Implements ICommand
        Protected ext As String
        Protected filter As String
        Protected uriStorage As String

        Public Sub New(ByVal control As IRichEditControl)
            MyBase.New(control)
        End Sub

        Protected Overrides Sub ExecuteCore()
            SaveFile(ext, filter)
        End Sub

        Public Overrides Sub UpdateUIState(ByVal state As DevExpress.Utils.Commands.ICommandUIState)
            MyBase.UpdateUIState(state)
            ' Disable a command if there is nothing to save
            If state.Enabled Then
                state.Enabled = Control.Document.Selection.Length > 0
            End If
        End Sub

        ' Provide default glyph for a command item
        Public Overridable ReadOnly Property ImageSource() As ImageSource
            Get
                Return MyBase.Image.Source
            End Get
        End Property

        ' Each descendant has its own implementation
        Protected MustOverride Function GetSelectedContents() As Byte()

        Private Sub SaveFile(ByVal ext As String, ByVal filter As String)
            Dim sfDialog As SaveFileDialog = New SaveFileDialog()

            sfDialog.DefaultExt = ext
            sfDialog.Filter = filter
            sfDialog.FilterIndex = 1

            If sfDialog.ShowDialog() = True Then
                Dim sm As Stream = sfDialog.OpenFile()
                If sm IsNot Nothing Then
                        uriStorage = sfDialog.SafeFileName
                        Dim bytes As Byte() = GetSelectedContents()
                        sm.Write(bytes, 0, bytes.Length)
                        sm.Close()
                End If
            End If
        End Sub

        ' Required for DXBars support
        #Region "ICommand Members"

        Private Overloads Function CanExecute(ByVal parameter As Object) As Boolean Implements ICommand.CanExecute
            Return CanExecute()
        End Function

        Private Custom Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
            AddHandler(ByVal value As EventHandler)
            End AddHandler
            RemoveHandler(ByVal value As EventHandler)
            End RemoveHandler
            RaiseEvent(ByVal sender As System.Object, ByVal e As System.EventArgs)
            End RaiseEvent
        End Event

        Private Overloads Sub Execute(ByVal parameter As Object) Implements ICommand.Execute
            Execute()
        End Sub

#End Region
    End Class
    #End Region ' #basecommand

    #Region "#openxmlcommand"
    Friend Class MyFileSaveAsDocxCommand
        Inherits MyFileSaveAsCommand

        Public Sub New(ByVal control As IRichEditControl)
            MyBase.New(control)
            ext = "docx"
            filter = "Word 2007 files (*.docx)|*.docx|All files (*.*)|*.*"
        End Sub

        Protected Overrides Sub ExecuteCore()
            MyBase.ExecuteCore()
        End Sub

        Protected Overrides Function GetSelectedContents() As Byte()
            Return Control.Document.GetOpenXmlBytes(Control.Document.Selection)
        End Function

        Public Overrides ReadOnly Property MenuCaption() As String
            Get
                Return "Save As DOCX"
            End Get
        End Property
    End Class
    #End Region ' #openxmlcommand

    Friend Class MyFileSaveAsRtfCommand
        Inherits MyFileSaveAsCommand

        Public Sub New(ByVal control As IRichEditControl)
            MyBase.New(control)
            ext = "rtf"
            filter = "Rich Text Format (*.rtf)|*.rtf|All files (*.*)|*.*"
        End Sub

        Protected Overrides Sub ExecuteCore()
            MyBase.ExecuteCore()
        End Sub

        Protected Overrides Function GetSelectedContents() As Byte()
            Dim s As String = Control.Document.GetRtfText(Control.Document.Selection)
            Dim encoding As System.Text.UTF8Encoding = New System.Text.UTF8Encoding()
            Return encoding.GetBytes(s)
        End Function

        Public Overrides ReadOnly Property MenuCaption() As String
            Get
                Return "Save As RTF"
            End Get
        End Property

    End Class


    Friend Class MyFileSaveAsHtmlEmbedCommand
        Inherits MyFileSaveAsCommand
        Public Sub New(ByVal control As IRichEditControl)
            MyBase.New(control)
            ext = "html"
            filter = "Hypertext Markup Language(*.html)|*.html|All files (*.*)|*.*"
        End Sub

        Protected Overrides Sub ExecuteCore()
            MyBase.ExecuteCore()
        End Sub

        Protected Overrides Function GetSelectedContents() As Byte()
            ' Embed images base64 encoded
            Dim canEmbed As Boolean = Control.InnerControl.Options.Export.Html.EmbedImages
            Control.InnerControl.Options.Export.Html.EmbedImages = True
            Dim s As String = Control.Document.GetHtmlText(Control.Document.Selection, Nothing)
            Control.InnerControl.Options.Export.Html.EmbedImages = canEmbed
            Dim encoding As System.Text.UTF8Encoding = New System.Text.UTF8Encoding()
            Return encoding.GetBytes(s)
        End Function

        Public Overrides ReadOnly Property MenuCaption() As String
            Get
                Return "Save As HTML (images embedded)"
            End Get
        End Property

    End Class

    #Region "#htmlsimplecommand"
    Friend Class MyFileSaveAsHtmlExternalCommand
        Inherits MyFileSaveAsCommand
        Private provider As MyUriProvider = Nothing

        Public Sub New(ByVal control As IRichEditControl)
            MyBase.New(control)
        End Sub

        Public Overrides Sub UpdateUIState(ByVal state As DevExpress.Utils.Commands.ICommandUIState)
            MyBase.UpdateUIState(state)
            ' Check for access to special folders
            If state.Enabled Then
                state.Enabled = Application.Current.HasElevatedPermissions
            End If
        End Sub

        Protected Overrides Sub ExecuteCore()

            Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            provider = New MyUriProvider(path)
            Dim bytes As Byte() = GetSelectedContents()
            Dim fileName As String = String.Format("{0}\{1}.html", path, Guid.NewGuid())
            Using stream As FileStream = New FileStream(fileName, FileMode.Create, FileAccess.Write)
                If stream IsNot Nothing Then
                    stream.Write(bytes, 0, bytes.Length)
                    MessageBox.Show("HTML snippet is saved to " & fileName)
                    stream.Close()
                End If
            End Using
        End Sub

        Protected Overrides Function GetSelectedContents() As Byte()
            Dim s As String = Control.Document.GetHtmlText(Control.Document.Selection, provider)
            Dim encoding As System.Text.UTF8Encoding = New System.Text.UTF8Encoding()
            Return encoding.GetBytes(s)
        End Function

        Public Overrides ReadOnly Property MenuCaption() As String
            Get
                Return "Save HTML to MyDocuments"
            End Get
        End Property
    End Class
       #End Region ' #htmlsimplecommand
End Namespace
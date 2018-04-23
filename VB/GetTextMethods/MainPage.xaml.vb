Imports Microsoft.VisualBasic
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
#Region "#usingsinmain"
Imports System.Collections.Generic
Imports System.IO
Imports DevExpress.Xpf.Bars
Imports DevExpress.Xpf.RichEdit
Imports DevExpress.XtraRichEdit
Imports DevExpress.Xpf.RichEdit.Menu
Imports System.Windows.Input
#End Region ' #usingsinmain

Namespace GetTextMethods
    Partial Public Class MainPage
        Inherits UserControl
        Public Sub New()
            InitializeComponent()
            AddHandler richEditControl1.PopupMenuShowing, AddressOf richEditControl1_PopupMenuShowing
            AddHandler richEditControl1.Loaded, AddressOf richEditControl1_Loaded
        End Sub

        Private Sub richEditControl1_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim stream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("sample_document.docx")
            richEditControl1.LoadDocument(stream, DocumentFormat.OpenXml)
            FillCommandDictionary()
        End Sub


        #Region "#addmenu"
        Private _commands As Dictionary(Of String, Object) = New Dictionary(Of String, Object)()

        Private Sub FillCommandDictionary()
            _commands.Add("DOCX", New MyFileSaveAsDocxCommand(richEditControl1))
            _commands.Add("RTF", New MyFileSaveAsRtfCommand(richEditControl1))
            _commands.Add("HTMLEMB", New MyFileSaveAsHtmlEmbedCommand(richEditControl1))
            _commands.Add("HTMLEXT", New MyFileSaveAsHtmlExternalCommand(richEditControl1))
        End Sub

        Private Sub richEditControl1_PopupMenuShowing(ByVal sender As Object, ByVal e As PopupMenuShowingEventArgs)
            e.Menu.ItemLinks.Add(New BarItemLinkSeparator())

            For Each kvp As KeyValuePair(Of String,Object) In _commands
                Dim item1 As RichEditMenuItem = New RichEditMenuItem()
                Dim cmd As MyFileSaveAsCommand = CType(kvp.Value, MyFileSaveAsCommand)
                item1.Command = cmd
                item1.Content = cmd.MenuCaption
                item1.Glyph = cmd.ImageSource

                e.Menu.ItemLinks.Add(item1)
            Next kvp

        End Sub
#End Region ' #addmenu

    End Class
End Namespace

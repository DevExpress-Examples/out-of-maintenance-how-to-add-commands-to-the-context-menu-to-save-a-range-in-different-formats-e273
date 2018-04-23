using System.Reflection;
using System.Windows;
using System.Windows.Controls;
#region #usingsinmain
using System.Collections.Generic;
using System.IO;
using DevExpress.Xpf.Bars;
using DevExpress.Xpf.RichEdit;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.Menu;
using System.Windows.Input;
#endregion #usingsinmain

namespace GetTextMethods
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
            richEditControl1.PopupMenuShowing += new PopupMenuShowingEventHandler(richEditControl1_PopupMenuShowing);
            richEditControl1.Loaded += new RoutedEventHandler(richEditControl1_Loaded);
        }

        void richEditControl1_Loaded(object sender, RoutedEventArgs e)
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("GetTextMethods.sample_document.docx");
            richEditControl1.LoadDocument(stream, DocumentFormat.OpenXml);
            FillCommandDictionary();
        }


        #region #addmenu
        Dictionary<string, object> _commands = new Dictionary<string, object>();

        void FillCommandDictionary() 
        {
            _commands.Add("DOCX", new MyFileSaveAsDocxCommand(richEditControl1)); 
            _commands.Add("RTF", new MyFileSaveAsRtfCommand(richEditControl1));
            _commands.Add("HTMLEMB", new MyFileSaveAsHtmlEmbedCommand(richEditControl1));
            _commands.Add("HTMLEXT", new MyFileSaveAsHtmlExternalCommand(richEditControl1));
        }

        private void  richEditControl1_PopupMenuShowing(object sender, PopupMenuShowingEventArgs e)
        {
            e.Menu.ItemLinks.Add(new BarItemLinkSeparator());

            foreach (KeyValuePair<string,object> kvp in _commands)
            {
                RichEditMenuItem item1 = new RichEditMenuItem();
                MyFileSaveAsCommand cmd = (MyFileSaveAsCommand)kvp.Value;
                item1.Command = cmd;
                item1.Content = cmd.MenuCaption;
                item1.Glyph = cmd.ImageSource;

                e.Menu.ItemLinks.Add(item1);
            }
        }
        #endregion #addmenu

    }
}

#region #usingsincommands
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.Commands;
using DevExpress.XtraRichEdit.Export.Html;
using System.Windows.Media;
#endregion #usingsincommands
namespace GetTextMethods
{
    #region #basecommand
    abstract class MyFileSaveAsCommand : SaveDocumentAsCommand, ICommand
    {
        protected string ext;
        protected string filter;
        protected string uriStorage;

        public MyFileSaveAsCommand(IRichEditControl control) : base(control) {}

        protected override void ExecuteCore()
        {
            SaveFile(ext, filter);
        }

        public override void UpdateUIState(DevExpress.Utils.Commands.ICommandUIState state)
        {
            base.UpdateUIState(state);
            // Disable a command if there is nothing to save
            if (state.Enabled) state.Enabled = Control.Document.Selection.Length > 0;
        }

        // Provide default glyph for a command item
        public virtual ImageSource ImageSource
        {
            get
            {
                return base.Image.Source;
            }
        }

        // Each descendant has its own implementation
        protected abstract byte[] GetSelectedContents();

        private void SaveFile(string ext, string filter)
        {
            SaveFileDialog sfDialog = new SaveFileDialog();

            sfDialog.DefaultExt = ext;
            sfDialog.Filter = filter;
            sfDialog.FilterIndex = 1;

            if (sfDialog.ShowDialog() == true)
            {
                Stream sm = sfDialog.OpenFile();    
                if (sm != null)
                    {
                        uriStorage = sfDialog.SafeFileName;
                        byte[] bytes = GetSelectedContents();
                        sm.Write(bytes, 0, bytes.Length);
                        sm.Close();
                    }
            }
        }

        // Required for DXBars support
        #region ICommand Members

        bool ICommand.CanExecute(object parameter)
        {
            return CanExecute();
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add {  }
            remove {  }
        }

        void ICommand.Execute(object parameter)
        {
            Execute();
        }

        #endregion
    }
    #endregion #basecommand

    #region #openxmlcommand
    class MyFileSaveAsDocxCommand : MyFileSaveAsCommand {

        public MyFileSaveAsDocxCommand(IRichEditControl control) : base(control) 
        {
            ext = "docx";
            filter = "Word 2007 files (*.docx)|*.docx|All files (*.*)|*.*";
        }

        protected override void ExecuteCore()
        {
            base.ExecuteCore();
        }
        
        protected override byte[] GetSelectedContents() 
        {
            return Control.Document.GetOpenXmlBytes(Control.Document.Selection);
        }

        public override string MenuCaption
        {
            get
            {
                return "Save As DOCX";
            }
        }
    }
    #endregion #openxmlcommand

    class MyFileSaveAsRtfCommand : MyFileSaveAsCommand
    {

        public MyFileSaveAsRtfCommand(IRichEditControl control)
            : base(control)
        {
            ext = "rtf";
            filter = "Rich Text Format (*.rtf)|*.rtf|All files (*.*)|*.*";
        }

        protected override void ExecuteCore()
        {
            base.ExecuteCore();
        }

        protected override byte[] GetSelectedContents()
        {
            string s = Control.Document.GetRtfText(Control.Document.Selection);
            System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
            return encoding.GetBytes(s);
        }

        public override string MenuCaption
        {
            get
            {
                return "Save As RTF";
            }
        }

    }

    #region #htmlembedcommand
    class MyFileSaveAsHtmlEmbedCommand : MyFileSaveAsCommand
    {
        public MyFileSaveAsHtmlEmbedCommand(IRichEditControl control)
            : base(control)
        {
            ext = "html";
            filter = "Hypertext Markup Language(*.html)|*.html|All files (*.*)|*.*";
        }

        protected override void ExecuteCore()
        {
            base.ExecuteCore();
        }

        protected override byte[] GetSelectedContents()
        {
            // Embed images base64 encoded
            bool canEmbed = Control.InnerControl.Options.Export.Html.EmbedImages;
            Control.InnerControl.Options.Export.Html.EmbedImages = true;
            string s = Control.Document.GetHtmlText(Control.Document.Selection, null);
            Control.InnerControl.Options.Export.Html.EmbedImages = canEmbed;
            System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
            return encoding.GetBytes(s);
        }

        public override string MenuCaption
        {
            get
            {
                return "Save As HTML (images embedded)";
            }
        }

    }
    #endregion #htmlembedcommand

    #region #htmlsimplecommand
    class MyFileSaveAsHtmlExternalCommand : MyFileSaveAsCommand
    {
        MyUriProvider provider = null;

        public MyFileSaveAsHtmlExternalCommand(IRichEditControl control)
            : base(control)
        {
        }

        public override void UpdateUIState(DevExpress.Utils.Commands.ICommandUIState state)
        {
            base.UpdateUIState(state);
            // Check for access to special folders
            if (state.Enabled) state.Enabled = Application.Current.HasElevatedPermissions;
        }
        
        protected override void ExecuteCore()
        {

            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            provider = new MyUriProvider(path);
            byte[] bytes = GetSelectedContents();
            string fileName = String.Format("{0}\\{1}.html", path, Guid.NewGuid());
            using (FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                if (stream != null)
                {
                    stream.Write(bytes, 0, bytes.Length);
                    MessageBox.Show("HTML snippet is saved to " + fileName);
                    stream.Close();
                }
            }
        }

        protected override byte[] GetSelectedContents()
        {
            string s = Control.Document.GetHtmlText(Control.Document.Selection, provider);
            System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
            return encoding.GetBytes(s);
        }

        public override string MenuCaption
        {
            get
            {
                return "Save HTML to MyDocuments";
            }
        }
    }
       #endregion #htmlsimplecommand
}
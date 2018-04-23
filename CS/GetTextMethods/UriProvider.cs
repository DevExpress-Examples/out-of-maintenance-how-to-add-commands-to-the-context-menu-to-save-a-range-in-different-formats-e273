#region #usingsinuriprovider
using System;
using System.IO;
using DevExpress.XtraRichEdit.Services;
using DevExpress.XtraRichEdit.Utils;
using DevExpress.Office.Utils;
#endregion #usingsinuriprovider
namespace GetTextMethods
{
    #region #uriprovider
    public class MyUriProvider : IUriProvider
    {
        string rootDirectory;
        public MyUriProvider(string rootDirectory)
        {
            if (String.IsNullOrEmpty(rootDirectory))
                Exceptions.ThrowArgumentException("rootDirectory", rootDirectory);
            this.rootDirectory = rootDirectory;
        }

        public string CreateCssUri(string rootUri, string styleText, string relativeUri)
        {
            // This method is not called when HTML content is obtained via the Document.GetHtmlText method.
            // Styles are placed within the <style> tag. 
            return String.Empty;
        }
        
        public string CreateImageUri(string rootUri, DevExpress.Office.Utils.OfficeImage image, string relativeUri)
        {
            string imagesDir = String.Format("{0}\\{1}", this.rootDirectory, rootUri.Trim('/'));
            if (!Directory.Exists(imagesDir))
                Directory.CreateDirectory(imagesDir);
            string imageName = String.Format("{0}\\{1}.png", imagesDir, Guid.NewGuid());
            byte[] bytes = image.GetImageBytesSafe(OfficeImageFormat.Png);
            using (FileStream stream = new FileStream(imageName, FileMode.Create, FileAccess.Write))
            {
                stream.Write(bytes, 0, bytes.Length);
                stream.Close();
            }
            return GetRelativePath(imageName);
        }
        
        string GetRelativePath(string path)
        {
            string substring = path.Substring(this.rootDirectory.Length);
            return substring.Replace("\\", "/").Trim('/');
        }
    }
    #endregion #uriprovider
}

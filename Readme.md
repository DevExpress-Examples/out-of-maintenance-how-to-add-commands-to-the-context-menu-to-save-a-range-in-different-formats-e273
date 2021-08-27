<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128606417/12.1.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E2731)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [GetTextMethodsTestPage.aspx](./CS/GetTextMethods.Web/GetTextMethodsTestPage.aspx) (VB: [GetTextMethodsTestPage.aspx](./VB/GetTextMethods.Web/GetTextMethodsTestPage.aspx))
* [Silverlight.js](./CS/GetTextMethods.Web/Silverlight.js)
* [Commands.cs](./CS/GetTextMethods/Commands.cs) (VB: [Commands.vb](./VB/GetTextMethods/Commands.vb))
* [MainPage.xaml](./CS/GetTextMethods/MainPage.xaml) (VB: [MainPage.xaml](./VB/GetTextMethods/MainPage.xaml))
* [MainPage.xaml.cs](./CS/GetTextMethods/MainPage.xaml.cs) (VB: [MainPage.xaml.vb](./VB/GetTextMethods/MainPage.xaml.vb))
* [UriProvider.cs](./CS/GetTextMethods/UriProvider.cs) (VB: [UriProvider.vb](./VB/GetTextMethods/UriProvider.vb))
<!-- default file list end -->
# How to add commands to the context menu to save a range in different formats 


<p>This example illustrates how you can create a set of commands which descend from a common base class and add them to context menu of the RichEditControl. Every command implements a certain method used to get selected range in specific format.<br />
To add a menu item, the <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraRichEditRichEditControl_PopupMenuShowingtopic"><u>RichEditControl.PopupMenuShowing event</u></a> is handled. A new <a href="http://documentation.devexpress.com/#Silverlight/clsDevExpressXtraRichEditMenuRichEditMenuItemtopic"><u>RichEditMenuItem</u></a> instance is created and associated with a command. The command provides <a href="http://documentation.devexpress.com/#Silverlight/DevExpressXpfBarsBarItem_Contenttopic"><u>Content</u></a> and <a href="http://documentation.devexpress.com/#Silverlight/DevExpressXpfBarsBarItem_Glyphtopic"><u>Glyph</u></a> for displaying the menu item.<br />
All commands descend from the <a href="http://documentation.devexpress.com/#Silverlight/clsDevExpressXtraRichEditCommandsSaveDocumentAsCommandtopic"><u>SaveDocumentAsCommand</u></a> command and implement the <a href="http://msdn.microsoft.com/en-us/library/system.windows.input.icommand.aspx"><u>ICommand</u></a> interface. The <strong>ExecuteCore </strong>method is overridden and invokes the custom <strong>Save As</strong><strong>...</strong> dialog. The <strong>GetSelectedContents</strong> method is implemented in command descendants to get a byte array of selected range in a specified format.</p><p>Since Silverlight application has quite limited file system access, the project is run in out-of-browser mode with elevated permissions. It enables the application to save data on the local host in special folders (e.g. My Documents). If the application is run without elevated permissions, only streams obtained via the SaveFileDialog are available, and the custom UriProvider will throw security exception. The command's <strong>UpdateUIState</strong><strong> </strong>method is overridden to check for elevated permissions and disable the HTML export command if this criteria is not met. Likewise, the base <strong>UpdateUIState </strong>method disables a command if the selection length is zero so there is nothing to export. </p><p>To get selected range in a required format, the commands use the <a href="http://documentation.devexpress.com/#CoreLibraries/DevExpressXtraRichEditAPINativeSubDocument_GetOpenXmlBytestopic"><u>GetOpenXmlBytes</u></a>, <a href="http://documentation.devexpress.com/#Silverlight/DevExpressXtraRichEditAPINativeSubDocument_GetRtfTexttopic"><u>GetRtfText</u></a> and <a href="http://documentation.devexpress.com/#CoreLibraries/DevExpressXtraRichEditAPINativeSubDocument_GetHtmlTexttopic"><u>GetHtmlText</u></a> methods. The latter method gives rise to two commands - one that works without an <a href="http://documentation.devexpress.com/#CoreLibraries/clsDevExpressXtraRichEditServicesIUriProvidertopic"><u>UriProvider</u></a> and another that uses a custom provider to save images. A command without an UriProvider sets the <a href="http://documentation.devexpress.com/#Silverlight/DevExpressXtraRichEditExportHtmlDocumentExporterOptions_EmbedImagestopic"><u>EmbedImages</u></a> option to <strong>true</strong> so images are base64 encoded and embedded into the HTML page.</p>

<br/>



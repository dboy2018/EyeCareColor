using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace EyeCareColor
{
    public partial class ThisAddIn
    {
        private readonly HashSet<string> fileHashSet = new HashSet<string>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Globals.Start();
        }

        private string GetDocumentStatus()
        {
            var doc = Application.ActiveDocument;
            if (!fileHashSet.Contains(doc.Name))
            {
                fileHashSet.Add(doc.Name);
                try
                {
                    if (string.IsNullOrWhiteSpace(doc.Path))
                    {
                        var docName = doc.Name;
                        return "new";
                    }
                    else
                    {
                        var docName = doc.Name;
                        return "open";
                    }
                }
                catch
                {
                    return "error";
                }
            }

            return "exists";
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }


        #region VSTO 生成的代码

        /// <summary>
        ///     设计器支持所需的方法 - 不要修改
        ///     使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;

            Application.DocumentOpen += Application_DocumentOpen;
            ((ApplicationEvents4_Event) Application).NewDocument += ThisAddIn_NewDocument;
            Application.WindowActivate += Application_WindowActivate;
        }

        private void Application_WindowActivate(Document Doc, Window Wn)
        {
            var documentStatus = GetDocumentStatus();
            if ("new".Equals(documentStatus)) Globals.Start();
        }

        private void ThisAddIn_NewDocument(Document Doc)
        {
            Globals.Start();
        }


        private void Application_DocumentOpen(Document Doc)
        {
            Globals.Start();
        }

        #endregion
    }
}
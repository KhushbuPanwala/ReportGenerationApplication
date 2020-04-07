using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Reflection;
using ReportGeneration;

using System.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Graph = Microsoft.Office.Interop.Graph;
using Core = Microsoft.Office.Core;

using Microsoft.Office.Interop.PowerPoint;
//using Microsoft.Office.Interop.Graph;
//using Microsoft.Office.Core;


using Microsoft.Office.Interop.Word;
using WordDoc = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace ReportDownloadApplication
{
    public partial class _Default : System.Web.UI.Page
    {
        string directoryName = "AuditReports";
        WordDoc.Document wordDocument;
        Presentation pptPresentation;

        protected void Page_Load(object sender, EventArgs e)
        {
            WordReport wordReport = new WordReport();
            wordDocument = wordReport.CreateWordDocument();

            PowerPointReport excelReport = new PowerPointReport();
            pptPresentation = excelReport.CreatePPTFile();

            SaveFile();

        }

        /// <summary>
        /// Save file in word, pdf and ppt formate
        /// </summary>
        private void SaveFile()
        {

            //Check directory exist or not
            bool exists = System.IO.Directory.Exists(Server.MapPath(directoryName));

            if (!exists)
                System.IO.Directory.CreateDirectory(Server.MapPath(directoryName));

            //Create Word file
            string fileName = DateTime.Now.Ticks.ToString();
            object wordFilePath = HttpContext.Current.Server.MapPath(directoryName).ToString() + "\\" + fileName + ".docx";
            wordDocument.SaveAs(ref wordFilePath);

            //Create PDF file
            object pdfPath = HttpContext.Current.Server.MapPath(directoryName).ToString() + "\\" + fileName + ".pdf";
            wordDocument.ExportAsFixedFormat(pdfPath.ToString(), WdExportFormat.wdExportFormatPDF);
            wordDocument.Close();

            //Create ppt file
            object pptFilePath = HttpContext.Current.Server.MapPath(directoryName).ToString() + "\\" + DateTime.Now.Ticks.ToString() + ".pptx";
            pptPresentation.SaveAs(pptFilePath.ToString(), PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Core.MsoTriState.msoTrue);
            pptPresentation.Close();

            //Download word file
            Response.ContentType = "Application/docx";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName + ".docx");
            Response.TransmitFile(wordFilePath.ToString());

            //Download pdf file
            //Response.ContentType = "Application/pdf";
            //Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName + ".pdf");
            //Response.TransmitFile(pdfPath.ToString());

            //Download ppt file
            //Response.ContentType = "Application/pptx";
            //Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName + ".pptx");
            //Response.TransmitFile(pdfPath.ToString());


            DeleteAllTempFile();

            //Response.End();

        }

        /// <summary>
        /// Delete all created temporary file
        /// </summary>
        private void DeleteAllTempFile()
        {
            string[] filePaths = Directory.GetFiles(Server.MapPath(directoryName));
            foreach (string filePath in filePaths)
                File.Delete(filePath);
        }
    }
}

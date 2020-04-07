using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Microsoft.Office.Interop.Word;
//using WordDoc = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.IO;

namespace ReportGeneration
{
    public class WordReport
    {
        const string fontName = "verdana";

        //Create an instance for word app  
        private Application WordApp = new Application();

        //Create a new document  
        Document document;

        //Create a missing variable for missing value  
        private object missing = System.Reflection.Missing.Value;
        private object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

        public object oRng;
        private Paragraph oPara;
        private Range wrdRng;
        private Table oTable;
        private InlineShape oChart;
        private InlineShape oImage;

        /// <summary>
        /// Create word document file
        /// </summary>
        /// <returns>Created document</returns>
        public Document CreateWordDocument()
        {
            //Create a new document  
            document = WordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            WordApp.Visible = false;

            CreateHeader(document);

            CreateFooter(document);

            CreateTitle(document, 24, "Heading 1", false);

            CreateUnOrderedList(document);

            CreateImage(document);

            //Insert another paragraph.
            CreatePargraph(document, 24, "This is a sentence of normal text. Now here is a table:");

            CreateTable(document, 3, 5);

            CreatePargraph(document, 24, "And here's another table:");

            CreateTable(document, 5, 2);


            CreateLineChart(document);

            CreateBarPaiChart(document);

            //Add text after the chart.
            wrdRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("Thank You");

            return document;

        }

        /// <summary>
        /// Create line chart
        /// </summary>
        /// <param name="document">Created document</param>
        private void CreateLineChart(Document document)
        {
            object oClassType = "MSGraph.Chart.8";
            wrdRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oChart = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            //Demonstrate use of late bound oChart and oChartApp objects to manipulate the chart object with MSGraph.
            object oChartApp = oChart.OLEFormat.Object.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, oChart.OLEFormat.Object, null);

            //Change the chart type to Line.
            object[] Parameters = new Object[1];
            Parameters[0] = 4; //xlLine = 4
            oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty, null, oChart.OLEFormat.Object, Parameters);

            //Update the chart image and quit MSGraph.
            oChartApp.GetType().InvokeMember("Update", BindingFlags.InvokeMethod, null, oChartApp, null);
            oChartApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, oChartApp, null);
            //If desired, you can proceed from here using the Microsoft Graph Object model on the oChart and oChartApp objects to make additional changes to the chart.

            //Set the width of the chart.
            oChart.Width = WordApp.InchesToPoints(6.25f);
            oChart.Height = WordApp.InchesToPoints(3.57f);
        }


        /// <summary>
        /// Create chart
        /// </summary>
        /// <param name="document">Created document</param>
        private void CreateBarPaiChart(Document document)
        {
            oRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oChart = document.Content.InlineShapes.AddChart(Office.XlChartType.xlBarClustered, ref oRng);
            oChart = document.Content.InlineShapes.AddChart(Office.XlChartType.xlPie, ref oRng);
        }


        #region Detail Data Section

        /// <summary>
        /// Create header
        /// </summary>
        /// <param name="document">Current document</param>
        /// <param name="SpaceAfter">Space after header</param>
        /// <param name="headingText">heading value</param>
        /// <param name="isRange">range to define start header</param>
        private void CreateTitle(Document document, int SpaceAfter, string headingText, bool isRange)
        {
            //Insert a paragraph at the beginning of the document.
            if (isRange)
            {
                oRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oPara = document.Content.Paragraphs.Add(ref oRng);
            }
            else
            {
                oPara = document.Content.Paragraphs.Add(ref missing);
            }
            oPara.Range.Text = headingText;
            oPara.Range.Font.Bold = 1;
            oPara.Range.Font.ColorIndex = WdColorIndex.wdBlue;
            oPara.Range.Font.Name = fontName;
            oPara.Format.SpaceAfter = SpaceAfter;
            oPara.Range.InsertParagraphAfter();
        }

        /// <summary>
        /// Create paragraph
        /// </summary>
        /// <param name="document">Current document</param>
        /// <param name="SpaceAfter">Space after paragraph</param>
        /// <param name="details">Detail of paragraph</param>
        private void CreatePargraph(Document document, int SpaceAfter, string details)
        {
            //Insert another paragraph.         
            oRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara = document.Content.Paragraphs.Add(ref oRng);
            oPara.Range.Text = details;
            oPara.Range.Font.Bold = 0;
            oPara.Range.Font.ColorIndex = WdColorIndex.wdBlack;
            oPara.Range.Font.Name = fontName;
            oPara.Format.SpaceAfter = SpaceAfter;
            oPara.Range.InsertParagraphAfter();
        }

        /// <summary>
        /// Create table
        /// </summary>
        /// <param name="document">Current document</param>
        /// <param name="noRows">no. of rows in table</param>
        /// <param name="noCols">no. of column in table</param>
        private void CreateTable(Document document, int noRows, int noCols)
        {

            wrdRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = document.Tables.Add(wrdRng, noRows, noCols, ref missing, ref missing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;

            oTable.Borders.Enable = 1;
            oTable.AllowAutoFit = true;
            //int row, cell;
            string strText;
            for (int row = 1; row <= noRows; row++)
                for (int cell = 1; cell <= noCols; cell++)
                {
                    strText = "r" + row + "c" + cell;
                    oTable.Cell(row, cell).Range.Text = strText;

                    oTable.Rows[1].Range.Font.Bold = 1;
                    oTable.Rows[1].Range.Font.Italic = 1;
                    //oTable.Rows[1].Range.Font.ColorIndex = WdColorIndex.wdGray25;
                    oTable.Rows[1].Range.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    oTable.Rows[1].Range.Font.Size = 10;

                    //Change width of columns 1 & 2
                    oTable.Columns[1].Width = WordApp.InchesToPoints(2);
                    oTable.Columns[2].Width = WordApp.InchesToPoints(3);
                }
        }


        /// <summary>
        /// create Ordered list
        /// </summary>
        private void CreateOrderedList()
        {
            Application app = new Application();
            app.Visible = true;
            Document doc = app.Documents.Add();

            Range range = doc.Range(0, 0);

            range.ListFormat.ApplyNumberDefault();
            range.Text = "Birinci";
            range.InsertParagraphAfter();
            ListTemplate listTemplate = range.ListFormat.ListTemplate;

            Range subRange = doc.Range(range.StoryLength - 1);

            subRange.ListFormat.ApplyBulletDefault();
            subRange.ListFormat.ListIndent();
            subRange.Text = "Alt Birinci";
            subRange.InsertParagraphAfter();
            ListTemplate sublistTemplate = subRange.ListFormat.ListTemplate;

            Range subRange2 = doc.Range(subRange.StoryLength - 1);
            subRange2.ListFormat.ApplyListTemplate(sublistTemplate);
            subRange2.ListFormat.ListIndent();
            subRange2.Text = "Alt İkinci";
            subRange2.InsertParagraphAfter();

            Range range2 = doc.Range(range.StoryLength - 1);
            range2.ListFormat.ApplyListTemplateWithLevel(listTemplate, true);
            WdContinue isContinue = range2.ListFormat.CanContinuePreviousList(listTemplate);
            range2.Text = "İkinci";
            range2.InsertParagraphAfter();

            Range range3 = doc.Range(range2.StoryLength - 1);
            range3.ListFormat.ApplyListTemplate(listTemplate);
            range3.Text = "Üçüncü";
            range3.InsertParagraphAfter();
        }

        /// <summary>
        /// Create un ordered list 
        /// </summary>
        /// <param name="document">Current document</param>
        private void CreateUnOrderedList(Document document)
        {

            //unordered list
            Paragraph assets = document.Content.Paragraphs.Add();

            assets.Range.ListFormat.ApplyBulletDefault();
            string[] bulletItems = new string[] { "One", "Two", "Three" };

            for (int i = 0; i < bulletItems.Length; i++)
            {
                string bulletItem = bulletItems[i];
                if (i < bulletItems.Length - 1)
                    bulletItem = bulletItem + "\n";
                assets.Range.InsertBefore(bulletItem);
            }



        }



        private void CreateImage(Document document)
        {
            //string imagePath = Server.MapPath("logo/venish.jpg");
            string imagePath = "D:\\Promact Infotech\\KPMG\\Demo Application\\ReportDownloadApplication\\ReportDownloadApplication\\logo\\venish.jpg";
            oImage = document.InlineShapes.AddPicture(imagePath, ref missing, ref missing, ref missing);
            oImage.Width = WordApp.InchesToPoints(6.25f);
            oImage.Height = WordApp.InchesToPoints(3.57f);
        }

        #endregion

        /// <summary>
        /// Used for page setup
        /// </summary>
        /// <param name="document">Current document</param>
        private void SetupPage(Document document)
        {
            //Set paper Size
            document.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
            //Set paper Format
            document.PageSetup.PageHeight = 855;
            document.PageSetup.PageWidth = 595;
        }


        #region Header & Footer Section
        /// <summary>
        /// Create header
        /// </summary>
        /// <param name="document">Current document</param>
        private void CreateHeader(Document document)
        {
            //Add header into the document  
            foreach (Section section in document.Sections)
            {
                //Get the header range and add the header details.  
                Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                headerRange.Font.ColorIndex = WdColorIndex.wdBlue;
                headerRange.Font.Size = 10;
                headerRange.Text = "KPMG Limited";
            }
        }


        /// <summary>
        /// Create footer
        /// </summary>
        /// <param name="document">Current document</param>
        private void CreateFooter(Document document)
        {

            //string pageNum = "1";
            //WordApp.Selection.GoTo( WdGoToItem.wdGoToPage,  WdGoToDirection.wdGoToNext, ref missing, pageNum);
            // Range rngPageNum = WordApp.Selection.Range;
            ////Insert Next Page section break so that numbering can start at 1
            //rngPageNum.InsertBreak( WdBreakType.wdSectionBreakNextPage);

            // Section currSec = document.Sections[rngPageNum.Sections[1].Index];
            // HeaderFooter ftr = currSec.Footers[ WdHeaderFooterIndex.wdHeaderFooterPrimary];

            ////So that the footer content doesn't propagate to the previous section    
            //ftr.LinkToPrevious = false;
            //ftr.PageNumbers.RestartNumberingAtSection = true;
            //ftr.PageNumbers.StartingNumber = 1;

            ////If the total pages should not be the total in the document, just the section
            ////use the field SectionPages instead of NumPages
            //object TotalPages = Microsoft.Office.Interop.Word.WdFieldType.wdFieldSectionPages;
            //object CurrentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;
            // Range rngCurrSecFooter = ftr.Range;
            //rngCurrSecFooter.Fields.Add(rngCurrSecFooter, ref CurrentPage, ref missing, false);


            ////rngCurrSecFooter.InsertAfter(" of ");
            ////rngCurrSecFooter.Collapse( WdCollapseDirection.wdCollapseEnd);
            ////rngCurrSecFooter.Fields.Add(rngCurrSecFooter, ref TotalPages, ref missing, false);


            //Add the footers into the document  
            foreach (Section wordSection in document.Sections)
            {
                //Get the footer range and add the footer details.  
                Range footerRange = wordSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                footerRange.Text = "Footer text goes here";
            }
        }

        #endregion
    }
}

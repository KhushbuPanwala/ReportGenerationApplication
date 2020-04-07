using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Graph;
using Microsoft.Office.Core;
//using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Graph = Microsoft.Office.Interop.Graph;
using System.Drawing;

namespace ReportGeneration
{
    public class PowerPointReport
    {
        Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
        Presentation pptPresentation;
        CustomLayout customLayout;

        Slides slides;
        Slide slide;
        TextRange objText;

        public const string fontName = "Trebuchet MS";

        /// <summary>
        /// create ppt presantation file
        /// </summary>
        /// <returns></returns>
        public Presentation CreatePPTFile()
        {
            pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            pptApplication.Visible = MsoTriState.msoTrue;
            customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

            //cover slide
            slides = pptPresentation.Slides;
            slide = slides.AddSlide(1, customLayout);

            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = 0xFF3300;

            CreateOrderedList();

            CreateUnOrderList();

            //Create image in slide 
            //CreateImage();

            //Create title
            CreateTitle("Name of the client", 32);

            CreateTextBox("Internal Audit - Review for Audit Period_________", 32, 120, 120, 500, 80);

            //Create new slide
            CreateNewSlide(pptPresentation, customLayout);

            //Create image in slide 
            CreateChart();

            //Create new slide
            CreateNewSlide(pptPresentation, customLayout);

            CreateTable();

            return pptPresentation;

        }

        /// <summary>
        /// Create new slide 
        /// </summary>
        /// <param name="pptPresentation">current ppt presentation file</param>
        /// <param name="customLayout">custom layout for slide</param>
        private void CreateNewSlide(Presentation pptPresentation, CustomLayout customLayout)
        {
            // Create new Slide
            slides = pptPresentation.Slides;
            slide = slides.AddSlide(slide.SlideIndex + 1, customLayout);

            Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 120, 500, 200, 70);
            shape.TextFrame.TextRange.Text = "Footer text goes here" + slide.SlideIndex.ToString();
        }

        /// <summary>
        /// Create Ordered list
        /// </summary>
        private void CreateOrderedList()
        {
            // Add title
            slide.Shapes[1].TextFrame.TextRange.Text = "Title of slide.com";

            // Add items to list
            var bulletedList = slide.Shapes[2]; // Bulleted point shape
            var listTextRange = bulletedList.TextFrame.TextRange;
            listTextRange.Text = "Content goes here\nYou can add text\nItem 3";
            // Change the bullet character
            var format = listTextRange.Paragraphs().ParagraphFormat;
            //.Bullet.Number;
            //.Paragraphs[1].BulletType = TextBulletType.Numbered;

            format.Bullet.Character = (char)9675;
        }

        /// <summary>
        /// Create unordered list
        /// </summary>
        private void CreateUnOrderList()
        {
            // Add title
            slide.Shapes[1].TextFrame.TextRange.Text = "Title of slide.com";

            // Add items to list
            var bulletedList = slide.Shapes[2]; // Bulleted point shape
            var listTextRange = bulletedList.TextFrame.TextRange;
            listTextRange.Text = "Content goes here\n You can add text\nItem 3";

            // Change the bullet character
            var format = listTextRange.Paragraphs().ParagraphFormat;
            format.Bullet.Character = (char)9675;
        }

        /// <summary>
        /// Create textbox for data content
        /// </summary>
        /// <param name="title">text value</param>
        /// <param name="fontSize">ize of font</param>
        /// <param name="left">left position</param>
        /// <param name="top">top position</param>
        /// <param name="width">text box width</param>
        /// <param name="height">text box height</param>
        private void CreateTextBox(string title, int fontSize, int left, int top, int width, int height)
        {
            Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            shape.TextFrame.TextRange.Text = title;
            shape.TextFrame.TextRange.Font.Name = fontName;
            shape.TextFrame.TextRange.Font.Size = fontSize;
        }

        /// <summary>
        /// Create title of slide
        /// </summary>
        /// <param name="title">value of title</param>
        /// <param name="fontSize">font size of title</param>
        private void CreateTitle(string title, int fontSize)
        {
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = title;
            objText.Font.Name = fontName;
            objText.Font.Size = fontSize;
        }

        //private void CreateImage()
        //{
        //    //Create image in slide
        //    string innerChartFilePath = Server.MapPath("logo/venish.jpg");
        //    float width = 100;
        //    float height = 100;
        //    slide.Shapes.AddPicture(innerChartFilePath, C.MsoTriState.msoFalseore, Microsoft.Office.Core.MsoTriState.msoTrue, 10, slide.Shapes[2].Top, width, height);
        //}


        /// <summary>
        /// Create chart 
        /// </summary>
        private void CreateChart()
        {
            TextRange textRange = slide.Shapes[1].TextFrame.TextRange;
            textRange.Text = "My Chart";
            textRange.Font.Name = fontName;
            textRange.Font.Size = 24;
            Graph.Chart objChart = (Graph.Chart)slide.Shapes.AddOLEObject(150, 150, 480, 320, "MSGraph.Chart.8", "",
                MsoTriState.msoFalse, "", 0, "", MsoTriState.msoFalse).OLEFormat.Object;

            objChart.ChartType = Graph.XlChartType.xl3DPie;
            objChart.Legend.Position = Graph.XlLegendPosition.xlLegendPositionBottom;
            objChart.HasTitle = true;
            objChart.ChartTitle.Text = "Sales for Black Programming & Assoc.";
        }

        /// <summary>
        /// Create table 
        /// </summary>
        private void CreateTable()
        {
            Table pptTable = slide.Shapes.AddTable(5, 2, 36, 138, 648, 294).Table;

            for (int i = 1; i <= pptTable.Rows.Count; i++)
            {
                for (int j = 1; j <= pptTable.Columns.Count; j++)
                {
                    pptTable.Cell(i, j).Shape.Fill.BackColor.RGB = 0xFF3300;
                    pptTable.Cell(i, j).Shape.TextFrame.TextRange.Font.Size = 12;
                    //pptTable.Cell(i, j).Shape.Line.BackColor.RGB = 0xFF3300;
                    pptTable.Cell(i, j).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

                    pptTable.Cell(i, j).Borders[PpBorderType.ppBorderLeft].DashStyle = MsoLineDashStyle.msoLineLongDashDot;
                    pptTable.Cell(i, j).Borders[PpBorderType.ppBorderLeft].ForeColor.RGB = 0xff00ff;
                    pptTable.Cell(i, j).Borders[PpBorderType.ppBorderLeft].Weight = 1.0f;

                    pptTable.Cell(i, j).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    pptTable.Cell(i, j).Shape.TextFrame.TextRange.Text = string.Format("[{0},{1}]", i, j);
                }

            }

        }

    }
}

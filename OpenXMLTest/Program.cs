// See https://aka.ms/new-console-template for more information


using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Xml;
using DocumentFormat.OpenXml.Validation;
using System.Diagnostics;
using Microsoft.VisualBasic;

internal static class Program
{
    public static void Main()
    {
        Console.WriteLine("Hello, World!");

        using (var document = SpreadsheetDocument.Create("test.xlsx", SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sample Sheet"
            };

            sheets.Append(sheet);

            workbookPart.Workbook.Save();

            var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

            var row = new Row();
            row.AppendChild(new Cell()
            {
                CellValue = new CellValue("Hello, OpenXML!"),
                DataType = new EnumValue<CellValues>(CellValues.String)
            });

            sheetData.AppendChild(row);

            var xs = new[] { 0.0, 1.0, 2.0, 3.0, 4.0, 5.0 };
            var ys = new[] { 0.0, 1.0, 0.0, -1.0, 0.0, 1.0 };

            foreach(var data in xs.Zip(ys))
            {
                var x = data.First;
                var y = data.Second;
                var dataRow = new Row();
                dataRow.AppendChild(new Cell()
                {
                    CellValue = new CellValue(x),
                    DataType = new EnumValue<CellValues>(CellValues.Number)
                });
                dataRow.AppendChild(new Cell()
                {
                    CellValue = new CellValue(y),
                    DataType = new EnumValue<CellValues>(CellValues.Number)
                });

                sheetData.AppendChild(dataRow);
            }

            worksheetPart.Worksheet.Save();

            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            worksheetPart.Worksheet.Append(new Drawing()
            {
                Id = worksheetPart.GetIdOfPart(drawingsPart)
            });

            worksheetPart.Worksheet.Save();

            drawingsPart.WorksheetDrawing = new WorksheetDrawing();

            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "ja-JP" });

            var chart = chartPart.ChartSpace.AppendChild(
                new DocumentFormat.OpenXml.Drawing.Charts.Chart()
                );
            chart.AppendChild(new AutoTitleDeleted() { Val = true });

            var plotArea = chart.AppendChild(new PlotArea());
            var layout = plotArea.AppendChild(new Layout());

            var scatterChart = plotArea.AppendChild(
                new ScatterChart()
                {
                    ScatterStyle = new ScatterStyle() { Val = ScatterStyleValues.LineMarker }
                });

            var scatterChartSeries = scatterChart.AppendChild(
                new ScatterChartSeries()
                {
                    Index = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = 0 },
                    Order = new Order() { Val = 0 },
                    SeriesText = new SeriesText(new NumericValue() { Text = "Test" })
                });

            



            //var xAxis = scatterChartSeries.AppendChild(new ValueAxis()
            //{
            //   AxisId = new AxisId() { Val = new UInt32Value(100u)},
            //    AxisPosition = new AxisPosition() { Val = AxisPositionValues.Bottom }
            //});

            var xValues = scatterChartSeries.AppendChild(
                new DocumentFormat.OpenXml.Drawing.Charts.XValues()
                );

            var formulaX = "'Sample Sheet'!$A$2:$A$7";

            var numberReferenceX = xValues.AppendChild(new NumberReference()
            {
                Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaX }
            });

            var valueCacheX = numberReferenceX.AppendChild(new NumberingCache());
            valueCacheX.Append(new PointCount() { Val = (uint)xs.Length});
            var ix = 0;
            foreach (var x in xs)
            {
                valueCacheX.AppendChild(new NumericPoint()
                {
                    Index = (uint)ix++,
                }).Append(new NumericValue(x.ToString()));
            }
            // var numberCacheX = 

            //var yAxis = scatterChartSeries.AppendChild(new ValueAxis()
            //{
            //    AxisId = new AxisId() { Val = new UInt32Value(101u) },
            //    AxisPosition = new AxisPosition() { Val = AxisPositionValues.Left }
            //});

            var yValues = scatterChartSeries.AppendChild(
                new DocumentFormat.OpenXml.Drawing.Charts.YValues()
                );

            var formulaY = "'Sample Sheet!$B$2:$B$7";
            var numberReferenceY = yValues.AppendChild(new NumberReference()
            {
                Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaY }
            });
            var numberingCacheY = numberReferenceY.AppendChild(new NumberingCache());
            numberingCacheY.Append(new PointCount() { Val = (uint)ys.Length });
            var iy = 0;
            foreach (var y in ys)
            {
                numberingCacheY.AppendChild(new NumericPoint() {
                    Index = (uint)iy++
                }).Append(new NumericValue(y.ToString()));
            }
            scatterChart.Append(new AxisId()
            {
                Val = new UInt32Value(100u)
            });
            scatterChart.Append(new AxisId()
            {
                Val = new UInt32Value(101u)
            });

            plotArea.AppendChild(
                new ValueAxis(
                    new AxisId() { Val = new UInt32Value(100u) },
                    new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                    new Delete() { Val = false },
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                    new CrossingAxis() { Val = 101u },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) }
                    ));

            plotArea.AppendChild(
                new ValueAxis(
                    new AxisId() { Val = new UInt32Value(101u) },
                    new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                    new Delete() { Val = false },
                    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                    new CrossingAxis() {  Val = 100u },
                    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero)}
                    ));

            chart.Append(
                new PlotVisibleOnly() { Val = true }
                );

            chartPart.ChartSpace.Save();

            var twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                new ColumnId("0"),
                new ColumnOffset("0"),
                new RowId("2"),
                new RowOffset("0")
                ));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                new ColumnId("8"),
                new ColumnOffset("0"),
                new RowId("12"),
                new RowOffset("0")
                ));

            var graphicFrame = twoCellAnchor.AppendChild(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";

            graphicFrame.Append(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties()
                    {
                        Id = 2u,
                        Name = "Sample Chart"
                    },
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()
                ));

            graphicFrame.Append(new Transform(
                new Offset() { X = 0L, Y = 0L },
                new Extents() { Cx = 0L, Cy = 0L }
                ));

            graphicFrame.Append(new Graphic(
                new DocumentFormat.OpenXml.Drawing.GraphicData(
                    new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }
                    )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                ));

            twoCellAnchor.Append(new ClientData());

            drawingsPart.WorksheetDrawing.Save();

            worksheetPart.Worksheet.Save();
        }

        using (var document = SpreadsheetDocument.Open("test.xlsx", true))
        {
            var validator = new OpenXmlValidator();
            int count = 0;
            foreach(var error in validator.Validate(document))
            {
                count++;
                Console.WriteLine("Error " + count);
                Console.WriteLine("Description: " + error.Description);
                Console.WriteLine("ErrorType: " + error.ErrorType);
                Console.WriteLine("Node: " + error.Node);
                Console.WriteLine("Path: " + error.Path.XPath);
                Console.WriteLine("Part: " + error.Part.Uri);
                Console.WriteLine("-------------------------------------------");
            }
            Console.WriteLine("count={0}", count);
        }
    }
}


using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;

TemplateTesting();

 void TemplateTesting()
{
    string filePath = "C:\\Testing\\Template_Creation\\FinalEditTemplatePPT\\NewFolder\\Likeability_Template.pptx";

    List<Model> modelTestValues = new List<Model>();

    for (int i = 0; i < 25; i++)
    {
        Model model = new Model();

        model.Label = "Logo " + i.ToString();
        model.value1 = 4;
        model.value2 = 10;
        model.value3 = 12;
        modelTestValues.Add(model);
    }

    Dictionary<string, string[]> labelColorMapping = new Dictionary<string, string[]>
    {
       { "Logo 2", new string[] { "FF0000", "0000FF", "00FF00" } }  // Red for Value1, Blue for Value2, Green for Value3
    };

    EditExcelValue(filePath, modelTestValues, labelColorMapping);
}

 void EditExcelValue(string filePath, List<Model> modelTestValues, Dictionary<string, string[]> labelColorMapping)
{
    using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
    {
        SlidePart slidePart = presentationDocument.PresentationPart.SlideParts.First();

        ChartPart chartPart = slidePart.ChartParts.First();

        EmbeddedPackagePart embeddedPackagePart = chartPart.EmbeddedPackagePart;

        BarChart barChart = chartPart.ChartSpace.Descendants<BarChart>().First();

        if (embeddedPackagePart != null)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(embeddedPackagePart.GetStream(), true))
            {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();

                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                int excelDataIndex = 2;
                Cell cellVariableObj;

                //manipulate the excel data embedded
                foreach (var modelValue in modelTestValues)
                {
                    cellVariableObj = sheetData.Descendants<Cell>().FirstOrDefault(c => c.CellReference == "A" + excelDataIndex.ToString());
                    if (cellVariableObj != null)
                    {
                        cellVariableObj.CellValue = new CellValue(modelValue.Label);
                        cellVariableObj.DataType = new EnumValue<CellValues>(CellValues.String);
                    }

                    cellVariableObj = sheetData.Descendants<Cell>().FirstOrDefault(c => c.CellReference == "B" + excelDataIndex.ToString());
                    if (cellVariableObj != null)
                    {
                        cellVariableObj.CellValue = new CellValue(modelValue.value1);
                        cellVariableObj.DataType = new EnumValue<CellValues>(CellValues.Number);
                    }

                    cellVariableObj = sheetData.Descendants<Cell>().FirstOrDefault(c => c.CellReference == "C" + excelDataIndex.ToString());
                    if (cellVariableObj != null)
                    {
                        cellVariableObj.CellValue = new CellValue(modelValue.value2);
                        cellVariableObj.DataType = new EnumValue<CellValues>(CellValues.Number);
                    }

                    cellVariableObj = sheetData.Descendants<Cell>().FirstOrDefault(c => c.CellReference == "D" + excelDataIndex.ToString());
                    if (cellVariableObj != null)
                    {
                        cellVariableObj.CellValue = new CellValue(modelValue.value3);
                        cellVariableObj.DataType = new EnumValue<CellValues>(CellValues.Number);
                    }
                    excelDataIndex++;
                }

                //save chart
                chartPart.ChartSpace.Save();
                //save workbook
                worksheetPart.Worksheet.Save();

                //Modification of formula is required as per the changes in the excel data
                for (int rowIndex = modelTestValues.Count + 2; rowIndex <= 31; rowIndex++)
                {
                    Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                    if (row != null)
                    {
                        row.Remove();
                    }
                }

                // Columns to update
                char[] columns = { 'B', 'C', 'D' };
                int columnIndex = 0;

                foreach (DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries barchartSeries in barChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries>())
                {
                    if (columnIndex < columns.Length)
                    {

                        //UpdateSeriesColorsAndTextColor(barchartSeries,
                        //                               modelTestValues,
                        //                               labelColorMapping,
                        //                               columnIndex);

                        char column = columns[columnIndex];

                        NumberReference numberReference = barchartSeries.Elements<DocumentFormat.OpenXml.Drawing.Charts.Values>().First().Elements<NumberReference>().First();
                        numberReference.Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula(formulaChanges(column, modelTestValues.Count() + 1));

                        // get numerical data of chart
                        NumberingCache numberingCache = numberReference.Elements<NumberingCache>().First();

                        //set number of numering cache data points equal to model count
                        numberingCache.GetFirstChild<PointCount>().Val = (uint)modelTestValues.Count;

                        // clear all existing data points for new data entry
                        numberingCache.RemoveAllChildren<NumericPoint>();

                        int idx = 0;

                        foreach (var modelValue in modelTestValues)
                        {
                            double chartValue;

                            switch (column)
                            {
                                case 'B':
                                    chartValue = modelValue.value1;
                                    break;
                                case 'C':
                                    chartValue = modelValue.value2;
                                    break;
                                case 'D':
                                    chartValue = modelValue.value3;
                                    break;
                                default:
                                    chartValue = 0;
                                    break;
                            }

                            numberingCache.Append(new NumericPoint()
                            {
                                Index = (uint)idx,
                                NumericValue = new DocumentFormat.OpenXml.Drawing.Charts.NumericValue(chartValue.ToString()),
                            });

                            if (chartValue == 0)
                            {
                                DocumentFormat.OpenXml.Drawing.Charts.DataLabels dataLabels = barchartSeries.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>();

                                DocumentFormat.OpenXml.Drawing.Charts.DataLabel dataLabel = new DocumentFormat.OpenXml.Drawing.Charts.DataLabel();
                                dataLabel.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = (uint)idx });
                                dataLabel.Append(new Delete() { Val = true });
                                dataLabels.Append(dataLabel);
                            }

                            idx++;
                        }
                        columnIndex++;
                    }

                    StringReference stringReference = barchartSeries.Elements<CategoryAxisData>().First().Elements<StringReference>().First();
                    stringReference.Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula("Sheet1!$A$2:$A$" + (modelTestValues.Count + 1).ToString());

                    StringCache stringCache = stringReference.Elements<StringCache>().First();

                    stringCache.GetFirstChild<PointCount>().Val = (uint)modelTestValues.Count;

                    stringCache.RemoveAllChildren<PointCount>();

                    int catIdx = 0;

                    foreach (var modelValue in modelTestValues)
                    {
                        stringCache.Append(new StringPoint() { Index = (uint)catIdx, NumericValue = new DocumentFormat.OpenXml.Drawing.Charts.NumericValue(modelValue.Label) });
                        StringPoint stringPoint = stringCache.Elements<StringPoint>().First();
                        catIdx++;
                    }
                }

                chartPart.ChartSpace.Save();

                //save worksheet
                worksheetPart.Worksheet.Save();

                //save slidePart
                slidePart.Slide.Save();
            }
        }
        presentationDocument.Save();
    }
}

void UpdateSeriesColorsAndTextColor(DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries barChartSeries,
                                         List<Model> modelTestValue,
                                         Dictionary<string, string[]> labelColorMapping,
                                         int seriesIndex)
{
    for (int pointIndex = 0; pointIndex < modelTestValue.Count; pointIndex++)
    {
        if (labelColorMapping.ContainsKey(modelTestValue[pointIndex].Label))
        {
            string[] colorArray = labelColorMapping[modelTestValue[pointIndex].Label];
            string color = seriesIndex < colorArray.Length ? colorArray[seriesIndex] : "FFFFFF"; // Default to white if index is out of bounds

            DocumentFormat.OpenXml.Drawing.Charts.DataPoint dataPoint = barChartSeries.Elements<DocumentFormat.OpenXml.Drawing.Charts.DataPoint>().ElementAtOrDefault(pointIndex);

            if (dataPoint == null)
            {
                dataPoint = new DocumentFormat.OpenXml.Drawing.Charts.DataPoint(new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = (uint)pointIndex });
                barChartSeries.Append(dataPoint);
            }

            ChartShapeProperties chartShapeProperties = dataPoint.Elements<ChartShapeProperties>().FirstOrDefault();
            if (chartShapeProperties == null)
            {
                chartShapeProperties = new ChartShapeProperties();
                dataPoint.Append(chartShapeProperties);
            }
            else
            {
                chartShapeProperties.RemoveAllChildren<SolidFill>();
            }

            SolidFill solidFill = new SolidFill(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = color });
            chartShapeProperties.Append(solidFill);
        }
    }
}

    string formulaChanges(char cell, int length)
{
    var f = $"Sheet1!${cell.ToString()}$2:${cell.ToString()}${length.ToString()}";
    return f;
}
public class Model
{
    public string Label { get; set; }
    public int value1 { get; set; }
    public int value2 { get; set; }
    public int value3 { get; set; }
}
    

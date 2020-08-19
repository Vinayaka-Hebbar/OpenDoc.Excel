using OfficeOpenXml;
using OpenDoc.Utils;
using OpenDoc.Xml;
using System;
using System.Drawing;

namespace OpenDoc.Excel
{
    internal sealed class ExcelChart
    {
        public readonly IRoot root;

        public ExcelChart(IRoot root)
        {
            this.root = root;
        }

        private void DrawScatter(XElement element, ExcelWorksheet worksheet, OfficeOpenXml.Drawing.Chart.ExcelScatterChart chart)
        {
            string legend = null;
            ExcelRange rangeX = null;
            ExcelRange rangeY = null;
            Color lineColor = Color.CadetBlue;
            Color markerColor = lineColor;
            OfficeOpenXml.Drawing.Chart.eMarkerStyle markerStyle = OfficeOpenXml.Drawing.Chart.eMarkerStyle.Square;
            foreach (var attr in element.Attributes(root.Namespace))
            {
                switch (attr.Name)
                {
                    case "legend":
                        legend = attr.Value;
                        break;
                    case "line-color":
                        lineColor = root.GetColor(attr.Value);
                        break;
                    case "marker-color":
                        markerColor = root.GetColor(attr.Value);
                        break;
                    case "x-min":
                        chart.XAxis.MinValue = XmlConvert.ToDouble(attr);
                        break;
                    case "x-max":
                        chart.XAxis.MaxValue = XmlConvert.ToDouble(attr);
                        break;
                    case "y-min":
                        chart.YAxis.MinValue = XmlConvert.ToDouble(attr);
                        break;
                    case "y-max":
                        chart.YAxis.MaxValue = XmlConvert.ToDouble(attr);
                        break;
                    case "marker":
                        markerStyle = (OfficeOpenXml.Drawing.Chart.eMarkerStyle)(int)Enum.Parse(typeof(Enums.MarkerType), attr.Value);
                        break;
                }
            }
            if (element.HasElements)
            {
                foreach (var node in element.Elements())
                {
                    switch (node.Name)
                    {
                        case "x-axis":
                            rangeX = GetRange(node, worksheet);
                            break;
                        case "y-axis":
                            rangeY = GetRange(node, worksheet);
                            break;

                    }
                }
            }
            var series = (OfficeOpenXml.Drawing.Chart.ExcelScatterChartSerie)chart.Series.Add(rangeX, rangeY);
            series.Border.Fill.Color = lineColor;
            series.Marker = markerStyle;
            series.MarkerColor = markerColor;
            series.MarkerLineColor = markerColor;
            series.Header = legend;
        }

        private void DrawLine(XElement element, ExcelWorksheet worksheet, OfficeOpenXml.Drawing.Chart.ExcelLineChart chart)
        {
            string legend = null;
            ExcelRange rangeX = null;
            ExcelRange rangeY = null;
            Color lineColor = Color.CadetBlue;
            Color markerColor = lineColor;
            OfficeOpenXml.Drawing.Chart.eMarkerStyle markerStyle = OfficeOpenXml.Drawing.Chart.eMarkerStyle.Square;
            foreach (var attr in element.Attributes(root.Namespace))
            {
                switch (attr.Name)
                {
                    case "legend":
                        legend = attr.Value;
                        break;
                    case "line-color":
                        lineColor = root.GetColor(attr.Value);
                        break;
                    case "marker-color":
                        markerColor = root.GetColor(attr.Value);
                        break;
                    case "x-min":
                        chart.XAxis.MinValue = XmlConvert.ToDouble(attr);
                        break;
                    case "x-max":
                        chart.XAxis.MaxValue = XmlConvert.ToDouble(attr);
                        break;
                    case "y-min":
                        chart.YAxis.MinValue = XmlConvert.ToDouble(attr);
                        break;
                    case "y-max":
                        chart.YAxis.MaxValue = XmlConvert.ToDouble(attr);
                        break;
                    case "marker":
                        markerStyle = (OfficeOpenXml.Drawing.Chart.eMarkerStyle)(int)Enum.Parse(typeof(Enums.MarkerType), attr.Value);
                        break;

                }
            }
            if (element.HasElements)
            {
                foreach (var node in element.Elements())
                {
                    switch (node.Name)
                    {
                        case "x-axis":
                            rangeX = GetRange(node, worksheet);
                            break;
                        case "y-axis":
                            rangeY = GetRange(node, worksheet);
                            break;

                    }
                }
            }
            var series = (OfficeOpenXml.Drawing.Chart.ExcelLineChartSerie)chart.Series.Add(rangeX, rangeY);
            series.Border.Fill.Color = lineColor;
            series.Marker = markerStyle;
            series.MarkerLineColor = markerColor;
            series.Header = legend;
        }

        private ExcelRange GetRange(XElement node, ExcelWorksheet worksheet)
        {
            int startRow = 1;
            int endRow = 1;
            int startColumn = 1;
            int endColumn = 1;
            foreach (var attr in node.Attributes(root.Namespace))
            {
                switch (attr.Name)
                {
                    case "start-row":
                        startRow = XmlConvert.ToInt32(attr);
                        break;
                    case "end-row":
                        endRow = XmlConvert.ToInt32(attr);
                        break;
                    case "start-column":
                        startColumn = XmlConvert.ToInt32(attr);
                        break;
                    case "end-column":
                        endColumn = XmlConvert.ToInt32(attr);
                        break;

                }
            }
            return worksheet.Cells[startRow, startColumn, endRow, endColumn];
        }

        internal OfficeOpenXml.Drawing.Chart.ExcelChart Create(IRangeContainer parent, XElement node, OfficeOpenXml.Drawing.Chart.eChartType type, string name)
        {
            var worksheet = parent.Sheet.WorkSheet;
            var chart = worksheet.Drawings.AddChart(name, type);
            foreach (var element in node.Elements(root.Namespace))
            {
                switch (type)
                {
                    case OfficeOpenXml.Drawing.Chart.eChartType.Line:
                    case OfficeOpenXml.Drawing.Chart.eChartType.LineStacked:
                    case OfficeOpenXml.Drawing.Chart.eChartType.LineMarkers:
                        DrawLine(element, worksheet, (OfficeOpenXml.Drawing.Chart.ExcelLineChart)chart);
                        break;
                    case OfficeOpenXml.Drawing.Chart.eChartType.XYScatter:
                    case OfficeOpenXml.Drawing.Chart.eChartType.XYScatterLines:
                    case OfficeOpenXml.Drawing.Chart.eChartType.XYScatterSmooth:
                        DrawScatter(element, worksheet, (OfficeOpenXml.Drawing.Chart.ExcelScatterChart)chart);
                        break;
                }
            }
            return chart;
        }
    }
}

using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace MedControl
{
    public static class ExcelHelper
    {
        public static List<string> LoadFirstColumn(string path)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return list;
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            var range = ws.RangeUsed();
            if (range == null) return list;
            var firstCol = range.FirstColumnUsed().ColumnNumber();
            var firstRow = range.FirstRowUsed().RowNumber();
            var lastRow = range.LastRowUsed().RowNumber();
            // assume header in first row; start from row+1 when possible
            var start = Math.Min(firstRow + 1, lastRow);
            for (int r = start; r <= lastRow; r++)
            {
                var val = ws.Cell(r, firstCol).GetString().Trim();
                if (!string.IsNullOrWhiteSpace(val)) list.Add(val);
            }
            // if nothing collected (maybe no header), try including the first row
            if (list.Count == 0)
            {
                for (int r = firstRow; r <= lastRow; r++)
                {
                    var val = ws.Cell(r, firstCol).GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(val)) list.Add(val);
                }
            }
            return list;
        }

        public static DataTable LoadToDataTable(string path)
        {
            var dt = new DataTable();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return dt;
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            var range = ws.RangeUsed();
            if (range == null) return dt;

            var firstRow = range.FirstRowUsed();
            // Create columns based on first row
            foreach (var cell in firstRow.Cells())
            {
                var header = cell.GetString();
                if (string.IsNullOrWhiteSpace(header)) header = $"Col{cell.Address.ColumnNumber}";
                if (!dt.Columns.Contains(header)) dt.Columns.Add(header);
                else dt.Columns.Add($"{header}_{cell.Address.ColumnNumber}");
            }
            // Data rows
            foreach (var row in range.RowsUsed().Skip(1))
            {
                var dr = dt.NewRow();
                int i = 0;
                foreach (var cell in row.Cells(1, dt.Columns.Count))
                {
                    dr[i++] = cell.GetValue<string>();
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public static void SaveDataTable(string path, DataTable table)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Dados");
            // headers
            for (int c = 0; c < table.Columns.Count; c++)
                ws.Cell(1, c + 1).Value = table.Columns[c].ColumnName;
            // rows
            for (int r = 0; r < table.Rows.Count; r++)
                for (int c = 0; c < table.Columns.Count; c++)
                    ws.Cell(r + 2, c + 1).Value = table.Rows[r][c]?.ToString();
            ws.Columns().AdjustToContents();
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            wb.SaveAs(path);
        }
    }
}

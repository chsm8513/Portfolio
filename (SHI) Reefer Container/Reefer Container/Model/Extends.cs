using Gem = GemBox.Spreadsheet;

namespace Reefer_Container
{
    public static class GemBoxHelper
    {
        public static string GetCellText(Gem.ExcelWorksheet worksheet, int row, int column)
        {
            Gem.ExcelCell cell = worksheet.Cells[row, column];
            if (cell?.Value == null)
                return string.Empty;

            try
            {
                return cell.GetFormattedValue()?.Trim() ?? string.Empty;
            }
            catch
            {
                return cell.Value?.ToString()?.Trim() ?? string.Empty;
            }
        }

        public static string GetCellTextResolveMerge(
            Gem.ExcelWorksheet worksheet,
            int row,
            int column)
        {
            Gem.ExcelCell cell = worksheet.Cells[row, column];
            if (cell == null)
                return string.Empty;

            if (cell.MergedRange != null)
            {
                Gem.ExcelCell topLeft = worksheet.Cells[
                    cell.MergedRange.FirstRowIndex,
                    cell.MergedRange.FirstColumnIndex];
                if (topLeft?.Value == null)
                    return string.Empty;

                try
                {
                    return topLeft.GetFormattedValue()?.Trim() ?? string.Empty;
                }
                catch
                {
                    return topLeft.Value?.ToString()?.Trim() ?? string.Empty;
                }
            }

            return GetCellText(worksheet, row, column);
        }
    }
}

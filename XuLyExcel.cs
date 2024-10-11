using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace test
{
    internal class XuLyExcel
    {
        private ExcelWorksheet workSheet;
        private readonly string excelFilePath;

        public XuLyExcel(string filePath)
        {
            excelFilePath = filePath;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            LoadExcelFile();
        }

        // Load the Excel file into memory
        private void LoadExcelFile()
        {
            if (File.Exists(excelFilePath))
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    workSheet = package.Workbook.Worksheets[0];
                }
            }
            else
            {
                throw new FileNotFoundException("Excel file not found.");
            }
        }

        // Tìm chuỗi trong cột
        public static int FindStringInColumn(ExcelWorksheet worksheet, string searchString, int startRow)
        {
            for (int row = startRow; row <= worksheet.Dimension.Rows; row++)
            {
                var cellValue = worksheet.Cells[row, 1].Value;
                if (cellValue != null && cellValue.ToString() == searchString)
                {
                    return row;
                }
            }
            return -1; // Không tìm thấy chuỗi
        }

        // Tìm chuỗi tên giảng viên
        public static string FindLecturerName(ExcelWorksheet worksheet, int startRow)
        {
            for (int row = startRow + 1; row <= worksheet.Dimension.Rows; row++)
            {
                var cellValue = worksheet.Cells[row, 1].Value?.ToString();
                if (!string.IsNullOrEmpty(cellValue))
                {
                    if (cellValue.StartsWith("Cán bộ giảng dạy: "))
                        return cellValue.Substring("Cán bộ giảng dạy: ".Length);
                    if (cellValue.StartsWith("CBGD "))
                        return cellValue.Substring("CBGD ".Length);
                }
            }
            return null;
        }

        // Extracts lecturer data between start and end markers and organizes it in a dictionary
        private Dictionary<string, StringBuilder> ExtractLecturerData(string startString, string endString)
        {
            var lecturerContents = new Dictionary<string, StringBuilder>();

            for (int row = 1; row <= workSheet.Dimension.Rows; row++)
            {
                var cellValue = workSheet.Cells[row, 1].Value?.ToString();
                if (cellValue == startString)
                {
                    int startRow = row;
                    int endRow = FindStringInColumn(workSheet, endString, startRow + 1);
                    if (endRow != -1)
                    {
                        string lecturerName = FindLecturerName(workSheet, startRow);
                        StringBuilder content = ExtractContentBetweenRows(startRow, endRow);

                        if (!lecturerContents.ContainsKey(lecturerName))
                            lecturerContents[lecturerName] = content;
                        else
                            lecturerContents[lecturerName].Append(content);
                    }
                }
            }
            return lecturerContents;
        }

        // Extracts content between two rows
        private StringBuilder ExtractContentBetweenRows(int startRow, int endRow)
        {
            var content = new StringBuilder();
            for (int i = startRow + 1; i < endRow; i++)
            {
                for (int col = 1; col <= workSheet.Dimension.Columns; col++)
                {
                    var cellValue = workSheet.Cells[i, col].Value;
                    if (cellValue != null)
                        content.AppendLine(cellValue.ToString());
                }
            }
            return content;
        }

        // Giải mã file PDF sang Excel thành văn bản có thể đọc bằng UnikeyNT
        public void EncodeData()
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var sourceWorksheet = package.Workbook.Worksheets[0];
                    int rowCount = sourceWorksheet.Dimension.Rows;
                    int colCount = sourceWorksheet.Dimension.Columns;

                    var excelApp = new Excel.Application();
                    excelApp.Visible = true;

                    // Bước 1: Copy dữ liệu từ Sheet 1
                    var excelData = new StringBuilder();
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = sourceWorksheet.Cells[row, col].Value;
                            if (cellValue != null)
                                excelData.Append(cellValue.ToString() + "\t");
                        }
                        excelData.AppendLine();
                    }

                    // Đặt dữ liệu vào clipboard
                    Clipboard.SetText(excelData.ToString());

                    // Bước 2: Thực hiện tổ hợp phím Ctrl+Shift+F9 để chuyển đổi dữ liệu
                    var workbook = excelApp.Workbooks.Add();
                    excelApp.SendKeys("^+{F9}");
                    Thread.Sleep(500); // Chờ để thực hiện chuyển đổi

                    string convertedData = Clipboard.GetText();
                    PasteConvertedDataToSheet(sourceWorksheet, convertedData);
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đã xảy ra lỗi: {ex.Message}");
            }
        }

        // Paste the converted data back into the worksheet
        private void PasteConvertedDataToSheet(ExcelWorksheet worksheet, string convertedData)
        {
            var rows = convertedData.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < rows.Length; i++)
            {
                var cells = rows[i].Split('\t');
                for (int j = 0; j < cells.Length; j++)
                    worksheet.Cells[i + 1, j + 1].Value = cells[j];
            }
        }

        // Phân chia file Excel
        public void SplitExcelFile(string inputFilePath)
        {
            string startString = "Phòng Đào Tạo";
            string endString = "Người lập biểu";

            if (!File.Exists(inputFilePath))
            {
                MessageBox.Show("Không tìm thấy tệp Excel.");
                return;
            }

            string excelFileName = Path.GetFileNameWithoutExtension(inputFilePath);
            var lecturerData = ExtractLecturerData(startString, endString);
            //SaveLecturerData(lecturerData, excelFileName);
        }

        // Tạo file TXT để lưu thông tin lịch dạy giảng viên
        //private void SaveLecturerData(Dictionary<string, StringBuilder> lecturerData, string excelFileName)
        //{
        //    string destinationFolder = Path.Combine(Path.GetDirectoryName(excelFilePath), excelFileName);

        //    if (!Directory.Exists(destinationFolder))
        //        Directory.CreateDirectory(destinationFolder);

        //    foreach (var kvp in lecturerData)
        //    {
        //        string fileName = $"{kvp.Key}.txt";
        //        string filePath = Path.Combine(destinationFolder, fileName);
        //        File.WriteAllText(filePath, kvp.Value.ToString());
        //    }
        //}
    }
}

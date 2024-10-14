using System;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using OfficeOpenXml;
using System.Threading;
using System.Collections.Generic;
using Path = System.IO.Path;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Text;

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

        public XuLyExcel()
        {

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
        private static int FindStringInColumn(ExcelWorksheet worksheet, string searchString, int startRow)
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
        private static string FindLecturerName(ExcelWorksheet worksheet, int startRow)
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

        public void ConvertPdfToExcel(string pdfFilePath, string excelFilePath)
        {
            try
            {
                using (PdfReader reader = new PdfReader(pdfFilePath))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                        if (worksheet != null)
                        {
                            int totalPage = reader.NumberOfPages; // tổng số trang Excel
                            int rowNum = 1;
                            for (int page = 1; page <= totalPage; page++)
                            {
                                string pageText = PdfTextExtractor.GetTextFromPage(reader, page);
                                string[] lines = pageText.Split('\n');
                                string extractedString = ""; // Khởi tạo biến để lưu trữ chuỗi được lấy từ dòng có độ dài 12
                                bool isLine12 = false; // Biến đánh dấu xem dòng hiện tại có độ dài là 12 hay không

                                foreach (string line in lines)
                                {
                                    string currentLine = line; // Tạo một biến mới để thay đổi nội dung của dòng

                                    // Kiểm tra độ dài của dòng
                                    if (currentLine.Length == 24)
                                    {
                                        // Lấy chuỗi có độ dài 12
                                        extractedString = currentLine.Replace(" ", "").Replace("-", "");
                                        isLine12 = true; // Đánh dấu dòng hiện tại là dòng có độ dài 12
                                    }
                                    else
                                    {
                                        // Nếu không phải là dòng có độ dài 12, kiểm tra vị trí có 2 dấu cách liên tiếp
                                        Match match = Regex.Match(currentLine, @"\d{2}\s{2}\w\s{1}|\d\s{2}\w\s{1}");
                                        if (match.Success)
                                        {
                                            // Chèn chuỗi có độ dài 12 vào sau vị trí có 2 dấu cách liên tiếp
                                            currentLine = currentLine.Insert(match.Index + match.Length, extractedString + " ");
                                        }
                                    }

                                    // Nếu dòng hiện tại không phải là dòng có độ dài 12, thêm vào worksheet
                                    if (!isLine12)
                                    {
                                        // Tiếp tục xử lý dòng như bình thường
                                        string[] data1 = currentLine.Split('?');
                                        for (int i = 0; i < data1.Length; i++)
                                        {
                                            int indexOfSecondSpace = data1[i].IndexOf(' ', data1[i].IndexOf(' ') + 1);
                                            // Kiểm tra xem có tồn tại dấu cách thứ hai không và trước dấu cách thứ hai là hai chữ số
                                            if (indexOfSecondSpace != -1 && Regex.IsMatch(data1[i].Substring(0, indexOfSecondSpace), @"\b\d{2}\b"))
                                            {
                                                // Kiểm tra ký tự đầu tiên sau dấu cách thứ hai có phải là chữ in hoa không
                                                if (Char.IsUpper(data1[i][indexOfSecondSpace + 1]))
                                                {
                                                    // Chèn ký tự '_' vào sau dấu cách thứ hai
                                                    data1[i] = data1[i].Insert(indexOfSecondSpace + 1, "_ ");
                                                }
                                            }
                                            worksheet.Cells[rowNum, i + 1].Value = data1[i];
                                        }
                                        rowNum++;
                                    }
                                    else
                                    {
                                        // Đặt lại biến đánh dấu để chuẩn bị cho dòng tiếp theo
                                        isLine12 = false;
                                    }

                                }
                            }
                        }
                        //Lưu tập tin Excel trước khi dán dữ liệu để tránh mất dữ liệu
                        FileInfo excelFile = new FileInfo(excelFilePath);
                        excelPackage.SaveAs(excelFile);
                    }
                }

                //Giải mã file lưu lại vào sheet 1
                encodeData(excelFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đã xảy ra lỗi convert excel: {ex.Message}");
            }
        }

        // Giải mã file PDF sang Excel thành văn bản có thể đọc bằng UnikeyNT
        private void encodeData(string excelFilePath)
        {
            try
            {
                // Load tập tin Excel
                using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(excelFilePath)))
                {
                    // Lấy dữ liệu từ Sheet 1
                    ExcelWorksheet sourceWorksheet = package.Workbook.Worksheets["Sheet1"];
                    int rowCount = sourceWorksheet.Dimension.Rows;
                    int colCount = sourceWorksheet.Dimension.Columns;

                    Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = false; // Hiển thị ứng dụng Excel


                    // Bước 1: Copy dữ liệu từ Sheet 1
                    string excelData = "";
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = sourceWorksheet.Cells[row, col].Value;
                            if (cellValue != null)
                            {

                                excelData += cellValue.ToString() + "\t";
                            }

                        }
                        excelData += "\n"; // Xuống dòng sau mỗi hàng
                    }

                    // Đặt dữ liệu vào clipboard
                    Clipboard.SetText(excelData);

                    // Bước 2: Thực hiện tổ hợp phím Ctrl+Shift+F9 để chuyển đổi dữ liệu
                    //SendKeys.SendWait("^+{F9}");

                    Excel.Workbook workbook = excelApp.Workbooks.Add();

                    excelApp.SendKeys("^+{F9}");

                    // Đóng workbook (nếu cần)
                    workbook.Close(SaveChanges: false);

                    // Đóng ứng dụng Excel
                    excelApp.Quit();


                    Thread.Sleep(500); // Chờ 0.5 giây để chuyển đổi được thực hiện

                    string convertedData = Clipboard.GetText();

                    // Bước 3: Dán dữ liệu đã chuyển đổi từ clipboard vào Sheet 2
                    string[] rows = convertedData.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < rows.Length; i++)
                    {
                        string[] cells = rows[i].Split('\t');
                        for (int j = 0; j < cells.Length; j++)
                        {
                            sourceWorksheet.Cells[i + 1, j + 1].Value = cells[j];

                        }
                    }
                    //Lưu tập tin Excel
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đã xảy ra lỗi: {ex.Message}");
            }
        }

        // Phân chia file Excel
        public void SplitExcelFile(string inputFilePath)
        {
            string startString = "Phòng Đào Tạo";
            string endString = "Người lập biểu";

            // Kiểm tra xem tệp Excel tồn tại
            if (!File.Exists(inputFilePath))
            {
                MessageBox.Show("Không tìm thấy tệp Excel.");
                return;
            }

            // Trích xuất tên của file Excel đã xuất ra
            string excelFileName = Path.GetFileNameWithoutExtension(inputFilePath);

            // Thiết lập giấy phép EPPlus
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Mở tệp Excel
            using (ExcelPackage package = new ExcelPackage(new FileInfo(inputFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Lấy sheet đầu tiên

                // Tạo danh sách thông tin của từng giảng viên
                Dictionary<string, StringBuilder> lecturerContents = new Dictionary<string, StringBuilder>();

                string currentLecturer = null;
                StringBuilder currentContent = null;


                // Tìm và ghi dữ liệu từng phần vào các file văn bản
                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    object cellValue = worksheet.Cells[row, 1].Value;
                    if (cellValue != null && cellValue.ToString() == startString)
                    {
                        // Tìm thấy chuỗi bắt đầu, tìm chuỗi kết thúc tương ứng
                        int startRow = row;
                        int endRow = FindStringInColumn(worksheet, endString, startRow + 1);

                        if (endRow != -1)
                        {
                            // Lưu nội dung của từng giảng viên vào từng StringBuilder
                            currentLecturer = FindLecturerName(worksheet, startRow);
                            currentContent = new StringBuilder();
                            for (int i = startRow + 1; i < endRow; i++)
                            {
                                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                                {
                                    object cellData = worksheet.Cells[i, col].Value;
                                    if (cellData != null)
                                    {
                                        currentContent.AppendLine(cellData.ToString());
                                    }
                                }
                            }

                            if (!lecturerContents.ContainsKey(currentLecturer))
                            {
                                lecturerContents.Add(currentLecturer, currentContent);

                            }
                            else
                            {
                                lecturerContents[currentLecturer].AppendLine();
                                lecturerContents[currentLecturer].Append(currentContent);
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Không tìm thấy chuỗi kết thúc sau dòng {startRow}.");
                        }
                    }
                }
            }
        }


    }
}

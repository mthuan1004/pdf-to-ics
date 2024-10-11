using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using OfficeOpenXml;
using System.Runtime.InteropServices;
using System.Threading;
using System.Collections.Generic;
using Path = System.IO.Path;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using ExcelDataReader;
using System.Globalization;
using System.Collections;

namespace test
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ToolTip tip = new ToolTip() { IsBalloon = true };
            tip.SetToolTip(btnChonFile, "Chọn file PDF cần chuyển đổi");
            tip.SetToolTip(btnCut, "Chuyển đổi ra ICS");
        }

        private void btnChonFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtPDFPath.Text = openFileDialog.FileName;
                }
            }
        }
        private void btnCut_Click(object sender, EventArgs e)
        {
            // Kiểm tra xem người dùng đã chọn file Excel chưa
            if (string.IsNullOrWhiteSpace(txtPDFPath.Text))
            {
                MessageBox.Show("Vui lòng chọn một file PDF trước khi tiếp tục.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string PDFFilePath = txtPDFPath.Text;
            string outputPDFFilePath = Path.ChangeExtension(PDFFilePath, ".xlsx");
            string outputDirectory = Path.GetDirectoryName(PDFFilePath);

            try
            {
                // Yêu cầu người dùng nhập tên thư mục
                string folderName = GetFolderNameFromUser();
                if (string.IsNullOrWhiteSpace(folderName))
                {
                    MessageBox.Show("Tên thư mục không được để trống.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Tạo thư mục mới để lưu các file ICS
                string icsFolder = Path.Combine(outputDirectory, folderName);
                Directory.CreateDirectory(icsFolder);

                // Thực hiện xử lý file Excel và lưu kết quả vào các file ICS
                ConvertPdfToExcel(PDFFilePath, outputPDFFilePath);
                SplitExcelFile(outputPDFFilePath);
                List<LichDayGiangVien> lichDayGiangViens = ProcessExcelData(outputPDFFilePath);
                GenerateIcsFile(lichDayGiangViens, icsFolder);

                LoadIcsFilesToDataGridView(icsFolder);

                // Hiển thị thông báo khi hoàn thành
                MessageBox.Show($"Danh sách lịch dạy đã được lưu vào thư mục '{folderName}'.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Xử lý các ngoại lệ và hiển thị thông báo lỗi
                MessageBox.Show($"Đã xảy ra lỗi khi xử lý file Excel: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Phương thức để yêu cầu người dùng nhập tên thư mục
        private string GetFolderNameFromUser()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "CHỌN NƠI LƯU TRỮ THƯ MỤC ICS:";
                DialogResult result = dialog.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    // Lấy đường dẫn thư mục đã chọn
                    string selectedFolderPath = dialog.SelectedPath;

                    // Kiểm tra xem thư mục có tồn tại không
                    if (Directory.Exists(selectedFolderPath))
                    {
                        return selectedFolderPath;
                    }
                    else
                    {
                        // Thư mục không tồn tại
                        MessageBox.Show("Thư mục không tồn tại.", "Lỗi");
                    }
                }

                return null;
            }
        }

        private void LoadIcsFilesToDataGridView(string folderPath)
        {
            // Xóa các hàng hiện tại trong DataGridView
            dataGridView1.Rows.Clear();

            // Lấy danh sách các tệp ICS trong thư mục
            string[] files = Directory.GetFiles(folderPath, "*.ics");

            // Thêm mỗi đường dẫn tệp vào DataGridView
            foreach (string file in files)
            {
                // Tạo một dòng mới
                DataGridViewRow row = new DataGridViewRow();

                // Tạo các ô mới, giá trị của ô là đường dẫn và tên tệp tin
                DataGridViewTextBoxCell cell = new DataGridViewTextBoxCell();
                DataGridViewTextBoxCell cell2 = new DataGridViewTextBoxCell();

                cell.Value = Path.GetFileName(file);
                cell2.Value = file;

                // Thêm các ô vào dòng
                row.Cells.Add(cell);
                row.Cells.Add(cell2);

                // Thêm dòng vào DataGridView
                dataGridView1.Rows.Add(row);
            }
        }



        private void ConvertPdfToExcel(string pdfFilePath, string excelFilePath)
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
                            int rowNum = 1;
                            for (int page = 1; page <= reader.NumberOfPages; page++)
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

        //Giải mã font PDF
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
                    excelApp.Visible = true; // Hiển thị ứng dụng Excel


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

        //Cắt file từng giảng viên
        private void SplitExcelFile(string inputFilePath)
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
                

                // Lưu đường dẫn của thư mục đích
                string destinationFolder = Path.Combine(Path.GetDirectoryName(inputFilePath), excelFileName);

                // Tạo thư mục mới nếu chưa tồn tại
                if (!Directory.Exists(destinationFolder))
                {
                    Directory.CreateDirectory(destinationFolder);
                }
                
                // Xuất ra các file văn bản
                foreach (var kvp in lecturerContents)
                {
                    string lecturerName = kvp.Key;
                    string fileName = $"{lecturerName}.txt";
                    string filePath = Path.Combine(destinationFolder, fileName);
                    File.WriteAllText(filePath, kvp.Value.ToString());
                }
            }
        }
        private bool IsSheetNameExists(ExcelPackage package, string sheetName)
        {
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                if (worksheet.Name == sheetName)
                {
                    return true;
                }
            }
            return false;
        }
        static int FindStringInColumn(ExcelWorksheet worksheet, string searchString, int startRow)
        {
            for (int row = startRow; row <= worksheet.Dimension.Rows; row++)
            {
                object cellValue = worksheet.Cells[row, 1].Value;
                if (cellValue != null && cellValue.ToString() == searchString)
                {
                    return row;
                }
            }
            return -1; // Không tìm thấy chuỗi
        }

        static string FindLecturerName(ExcelWorksheet worksheet, int startRow)
        {
            string lecturerName = null;

            for (int row = startRow + 1; row <= worksheet.Dimension.Rows; row++)
            {
                object cellValue = worksheet.Cells[row, 1].Value;
                if (cellValue != null)
                {
                    if (cellValue.ToString().StartsWith("Cán bộ giảng dạy: "))
                    {
                        lecturerName = cellValue.ToString().Substring("Cán bộ giảng dạy: ".Length);
                        break;
                    }
                    else if (cellValue.ToString().StartsWith("CBGD "))
                    {
                        lecturerName = cellValue.ToString().Substring("CBGD ".Length);
                        break;
                    }
                }
            }

            return lecturerName;
        }

        static List<LichDayGiangVien> ProcessExcelData(string filePath)
        {
            List<LichDayGiangVien> lichDayGiangViens = new List<LichDayGiangVien>();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var configuration = new ExcelReaderConfiguration
            {
                FallbackEncoding = Encoding.UTF8 // Thiết lập mã hóa UTF-8
            };
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream, configuration))
            {
                while (reader.Read())
                {
                    string cellValue = reader.GetString(0); // Đọc giá trị từ cột đầu tiên

                    if (cellValue.StartsWith("Cán bộ giảng dạy") || cellValue.StartsWith("CBGD"))
                    {
                        var giangVien = ProcessGiangVien(cellValue);
                        var lichDay = new List<MonHoc>();
                        // Duyệt qua các dòng tiếp theo cho đến khi gặp dòng "Thời gian học:"
                        while (reader.Read())
                        {
                            cellValue = reader.GetString(0);
                            if (cellValue.StartsWith("Thời gian học:"))
                            {
                                break;
                            }
                            else if (cellValue.StartsWith("CS"))
                            {
                                var monHoc = ProcessMonHoc(cellValue);
                                lichDay.Add(monHoc); // Thêm thông tin môn học vào danh sách lịch dạy
                            }
                        }

                        lichDayGiangViens.Add(new LichDayGiangVien(giangVien, lichDay)); // Thêm thông tin lịch dạy của giảng viên vào danh sách
                    }
                }
            }

            return lichDayGiangViens;
        }

        static GiangVien ProcessGiangVien(string cellValue)
        {
            string tenGiangVien = "";
            string maGiangVien = "";
            string[] separators = new string[] { "Cán bộ giảng dạy: ", "CBGD ", " (" };
            string[] result = cellValue.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            
            // Trích xuất thông tin
            List<string> lstr = result.ToList();
            if(lstr.Count() < 2)
            {
                tenGiangVien = cellValue;
            }
            else
            {
             tenGiangVien = result[0].Trim();
             maGiangVien = result[1].Trim(')');

            }

            return new GiangVien(maGiangVien, tenGiangVien, "sdt", "khoa");
        }

        static MonHoc ProcessMonHoc(string cellValue)
        {
            // Xử lý dòng để lấy thông tin về môn học
            // Ví dụ:
            string[] parts = cellValue.Split(' ');

            // Lấy mã môn học
            string maMonHoc = parts[0];

            // Lấy nhóm
            string nhom = parts[1];
           
            // Lấy thời gian
            string thoiGian = parts[parts.Length - 1];

            // Tách thời gian thành hai phần trước và sau dấu "-"
            string[] thoiGianParts = thoiGian.Split('-');
            string ngayBatDau = thoiGianParts[0];


            // Chuyển đổi chuỗi ngày bắt đầu sang đối tượng DateTime
            DateTime ngayBatDauDateTime = DateTime.ParseExact(ngayBatDau, "dd/MM/yy", CultureInfo.InvariantCulture);

            // Định dạng lại ngày theo định dạng yêu cầu
            string formattedNgayBatDau = ngayBatDauDateTime.ToString("yyyyMMddTHHmmssZ");




            string ngayKetThuc = thoiGianParts[1];
            // MessageBox.Show($"Ngày bắt đầu: {ngayBatDau}\nNgày kết thúc: {ngayKetThuc}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // Lấy phòng
            string phong = parts[parts.Length - 2];
            if (phong == "*")
            {
                return null;
            }

            // Lấy tiết học
            string tietHoc = parts[parts.Length - 3];

            // Lấy thứ
            string thu = parts[parts.Length - 4];

            // Lấy sĩ số
            string siSo = parts[parts.Length - 6];
            //int siSo = int.Parse(parts[parts.Length - 5]);

            // Lấy lớp
            string lop = parts[parts.Length - 7];

            // Lấy phần cuối cùng (tên môn học)
            string tenMonHoc = "";
            for (int i = parts.Length - 8; i > 2; i--)
            {
                if (!parts[i].StartsWith("CS") && !parts[i].Contains("/"))
                {
                    tenMonHoc = parts[i] + " " + tenMonHoc;
                }
                else
                {
                    break;
                }
            }
            tenMonHoc = tenMonHoc.Trim();

            // Trả về đối tượng MonHoc đã được xử lý
            return new MonHoc(maMonHoc, nhom, tenMonHoc, thoiGian, ngayBatDau, ngayKetThuc, phong, tietHoc, thu, siSo, lop);
        }

        static void GenerateIcsFile(List<LichDayGiangVien> lichDayGiangViens, string outputDirectory)
        {
            foreach (var lichDayGiangVien in lichDayGiangViens)
            {
                string giangVienFileName = Path.Combine(outputDirectory, $"{lichDayGiangVien.GiangVien.TenGiangVien}.ics");

                using (StreamWriter writer = new StreamWriter(giangVienFileName))
                {
                    writer.WriteLine("BEGIN:VCALENDAR");
                    writer.WriteLine("VERSION:2.0");
                    foreach (var monHoc in lichDayGiangVien.LichDay)
                    {
                        if(monHoc == null)
                        {
                            continue;
                        }
                        DateTime ngayBatDau = DateTime.ParseExact(monHoc.ThoiGianBD, "dd/MM/yy", CultureInfo.InvariantCulture);
                        DateTime ngayKetThuc = DateTime.ParseExact(monHoc.ThoiGianKT, "dd/MM/yy", CultureInfo.InvariantCulture);
                        int soNgay = (int)(ngayKetThuc - ngayBatDau).TotalDays; // Số ngày giữa ngày bắt đầu và ngày kết thúc
                        int soLanLapLai = (soNgay + 6) / 7; // Số lần lặp lại cần thiết để bao phủ tất cả các ngày trong tuần


                        writer.WriteLine("BEGIN:VEVENT");
                        writer.WriteLine($"DTSTART:{GetStartDateTime(monHoc.ThoiGianBD, monHoc.TietHoc, monHoc.Thu)}");
                        writer.WriteLine($"DTEND:{GetEndDateTime(monHoc.ThoiGianBD, monHoc.TietHoc, monHoc.Thu)}");
                        writer.WriteLine($"RRULE:FREQ=WEEKLY;WKST=MO;COUNT={soLanLapLai};BYDAY={GetDayOfWeek(monHoc.Thu)}");
                        writer.WriteLine($"DESCRIPTION:{monHoc.MaMonHoc}");
                        writer.WriteLine($"LOCATION:{monHoc.Phong}");
                        writer.WriteLine($"STATUS:CONFIRMED");
                        writer.WriteLine($"SUMMARY:[{monHoc.MaMonHoc}][{monHoc.Nhom}]{monHoc.TenMonHoc}");
                        writer.WriteLine("TRANSP:OPAQUE");
                        writer.WriteLine("END:VEVENT");

                    }
                    writer.WriteLine("END:VCALENDAR");

                }


            }

        }


        private static string GetStartDateTime(string NgayBatDau, string TietHoc, string thu)

        {

            DateTime ngayBatDauDateTime = DateTime.ParseExact(NgayBatDau, "dd/MM/yy", CultureInfo.InvariantCulture);

            // Chuyển đổi thứ thành kiểu DayOfWeek
            DayOfWeek thuEnum = (DayOfWeek)Enum.Parse(typeof(DayOfWeek), thu);

            DateTime startDateTime = ngayBatDauDateTime;
            int startDayOfWeek = (int)ngayBatDauDateTime.DayOfWeek;

            // Tìm thứ trong tuần của ngày bắt đầu
            int thuBatDau = (int)Enum.Parse(typeof(DayOfWeek), thu);

            // Tính sự chênh lệch giữa thứ của ngày bắt đầu và thứ của buổi học
            int dayDiff = thuBatDau - startDayOfWeek;
            if (dayDiff < 0) // Nếu thứ của buổi học nhỏ hơn thứ của ngày bắt đầu
            {
                dayDiff += 7; // Thêm 7 để đảm bảo kết quả là số dương
            }

            // Thêm sự chênh lệch vào ngày bắt đầu
            startDateTime = startDateTime.AddDays(dayDiff);
            startDateTime = startDateTime.AddDays(-1);

            // Xác định thời gian bắt đầu dựa vào giá trị của Tiết Học
            switch (TietHoc)
            {
                case "123":
                    startDateTime = startDateTime.AddHours(7).AddMinutes(00); // 07:30
                    break;
                case "456":
                    startDateTime = startDateTime.AddHours(9).AddMinutes(35); // 09:35
                    break;
                case "789":
                    startDateTime = startDateTime.AddHours(12).AddMinutes(35); // 12:35
                    break;
                case "012":
                    startDateTime = startDateTime.AddHours(15).AddMinutes(10); // 15:10
                    break;
                case "89012":
                    startDateTime = startDateTime.AddHours(13).AddMinutes(25); // 13:25
                    break;
                case "23456":
                    startDateTime = startDateTime.AddHours(7).AddMinutes(50); // 07:50
                    break;
                case "12345":
                case "123456":
                    startDateTime = startDateTime.AddHours(7).AddMinutes(00); // 07:00
                    break;
                case "789012":
                case "78901":
                    startDateTime = startDateTime.AddHours(12).AddMinutes(35); // 12:35
                    break;
                    // Các trường hợp khác...
            }

            // Trừ đi 1 ngày từ thời gian bắt đầu
            // startDateTime = startDateTime.AddDays(-1);

            // Chuyển đổi thời gian sang múi giờ GMT
            startDateTime = startDateTime.ToUniversalTime();

            // Format thời gian theo định dạng chuẩn cho Google Calendar
            string formattedDateTime = startDateTime.ToString("yyyyMMddTHHmmssZ");

            return formattedDateTime;

        }

        private static string GetEndDateTime(string NgayBatDau, string TietHoc, string thu)
        {

            DateTime ngayBatDauDateTime = DateTime.ParseExact(NgayBatDau, "dd/MM/yy", CultureInfo.InvariantCulture);

            // Chuyển đổi thứ thành kiểu DayOfWeek
            DayOfWeek thuEnum = (DayOfWeek)Enum.Parse(typeof(DayOfWeek), thu);
            DateTime startDateTime = ngayBatDauDateTime;
            int startDayOfWeek = (int)ngayBatDauDateTime.DayOfWeek;

            // Tìm thứ trong tuần của ngày bắt đầu
            int thuBatDau = (int)Enum.Parse(typeof(DayOfWeek), thu);

            // Tính sự chênh lệch giữa thứ của ngày bắt đầu và thứ của buổi học
            int dayDiff = thuBatDau - startDayOfWeek;
            if (dayDiff < 0) // Nếu thứ của buổi học nhỏ hơn thứ của ngày bắt đầu
            {
                dayDiff += 7; // Thêm 7 để đảm bảo kết quả là số dương
            }

            // Thêm sự chênh lệch vào ngày bắt đầu
            startDateTime = startDateTime.AddDays(dayDiff);
            startDateTime = startDateTime.AddDays(-1);
            // Xác định thời gian bắt đầu dựa vào giá trị của Tiết Học
            switch (TietHoc)
            {
                case "123":
                    startDateTime = startDateTime.AddHours(9).AddMinutes(30); // 07:30
                    break;
                case "456":
                    startDateTime = startDateTime.AddHours(12).AddMinutes(05); // 09:35
                    break;
                case "789":
                    startDateTime = startDateTime.AddHours(15).AddMinutes(05); // 09:35
                    break;
                case "012":
                    startDateTime = startDateTime.AddHours(17).AddMinutes(40); // 09:35
                    break;
                case "89012":
                    startDateTime = startDateTime.AddHours(17).AddMinutes(40); // 09:35
                    break;
                case "23456":
                    startDateTime = startDateTime.AddHours(12).AddMinutes(05); // 09:35
                    break;
                case "12345":
                    startDateTime = startDateTime.AddHours(11).AddMinutes(15); // 09:35
                    break;
                case "123456":
                    startDateTime = startDateTime.AddHours(12).AddMinutes(05); // 09:35
                    break;
                case "789012":
                    startDateTime = startDateTime.AddHours(17).AddMinutes(40); // 09:35
                    break;
                    
                case "78901":
                    startDateTime = startDateTime.AddHours(16).AddMinutes(50); // 09:35
                    break;
                    // Các trường hợp khác...
            }


            // Trừ đi 1 ngày từ thời gian bắt đầu
            //  startDateTime = startDateTime.AddDays(-1);

            // Chuyển đổi thời gian sang múi giờ GMT
            startDateTime = startDateTime.ToUniversalTime();

            // Format thời gian theo định dạng chuẩn cho Google Calendar
            string formattedDateTime = startDateTime.ToString("yyyyMMddTHHmmssZ");

            return formattedDateTime;
        }
        // Phương thức để lấy ngày trong tuần tiếp theo dựa trên một ngày và một ngày trong tuần
        private static string GetDayOfWeek(string thu)
        {
            switch (thu)
            {
                case "2":
                    return "MO";
                case "3":
                    return "TU";
                case "4":
                    return "WE";
                case "5":
                    return "TH";
                case "6":
                    return "FR";
                case "7":
                    return "SA";
                default:
                    // Xử lý trường hợp không xác định được ngày trong tuần, có thể làm gì đó ở đây
                    return "";
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 1 && dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {
                // Lấy giá trị của ô được double click
                string filePath = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                string folderPath = Path.GetDirectoryName(filePath);

                // Kiểm tra xem đường dẫn tồn tại trước khi mở nó
                if (File.Exists(filePath))
                {
                    // Mở đường dẫn tương ứng

                    Process.Start(folderPath);
                }
                else
                {
                    MessageBox.Show("Tệp không tồn tại.", "Lỗi");
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
    public class SheetInfo
    {
        public string LecturerName { get; set; }

        public SheetInfo(string lecturerName)
        {
            LecturerName = lecturerName;
        }
    }
}

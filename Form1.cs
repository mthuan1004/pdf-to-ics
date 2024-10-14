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
            XuLyExcel xulyExcel = new XuLyExcel();
            XuLyICS xulyICS = new XuLyICS();
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

                //Quy trình xử lý file từ PDF sang ICS 
                xulyExcel.ConvertPdfToExcel(PDFFilePath,outputPDFFilePath);
                xulyExcel.SplitExcelFile(outputPDFFilePath);

                List<LichDayGiangVien> lichDayGiangViens = LichDayGiangVien.ProcessExcelData(outputPDFFilePath);

                XuLyICS.GenerateIcsFile(lichDayGiangViens, icsFolder);

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
        //Hiển thị file & đường dẫn ICS trên DataGridView
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

        //Cắt file từng giảng viên
       
        //Kiểm tra Sheet GV tồn tại
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

        //Hàm tìm chuỗi trong dòng ( file Excel, KeyWork, dòng bắt đầu)
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

        //TÌm tên GV trong file Excel (file Excel, Dòng bắt đầu)
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

       

        // New method to handle processing of MonHoc including defaulting to "CS" if empty
        

       

    }
 
}

using iTextSharp.tool.xml.html;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    internal class XuLyFile
    {
        private string maMon = "E:\\PDF\\maMon.txt";
        private string log = "E:\\PDF\\log.txt";

        public XuLyFile(string path)
        {
            maMon = path;
        }

        public XuLyFile() { }

        //Đọc file maMon.txt
        public List<string> ReadCodesFromFile()
        {
            List<string> codeList = new List<string>();

            try
            {
                if (File.Exists(maMon))
                {
                    // Đọc tất cả các dòng từ file và lọc các dòng không phải là khoảng trắng
                    codeList = File.ReadAllLines(maMon)
                                   .Where(line => !string.IsNullOrWhiteSpace(line))
                                   .Select(line => line.Trim())
                                   .ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đọc mã môn thất bại: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return codeList;
        }


        // Ghi nối tiếp các file xử lý lỗi vào log
        public void AppendToFile(string content)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(log, true))
                {
                    writer.WriteLine(content);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi ghi thêm vào file: " + ex.Message);
            }
        }

    }
}

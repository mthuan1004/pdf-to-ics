using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test
{
    internal class DocFile
    {
        private string maMon;
        private string log;

        public DocFile(string path)
        {
            maMon = path;
        }

        public DocFile() { }

        //Đọc file maMon.txt
        public List<string> ReadCodesFromFile()
        {
            List<string> codeList = new List<string>();
            try
            {
                if (File.Exists(maMon))
                {
                    // Đọc tất cả các dòng từ file và lưu vào danh sách
                    string[] lines = File.ReadAllLines(maMon);
                    foreach (string line in lines)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            codeList.Add(line.Trim()); // Loại bỏ khoảng trắng và thêm vào danh sách
                        }
                    }
                }
                else
                {
                    Console.WriteLine("File không tồn tại.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi đọc file: " + ex.Message);
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
                Console.WriteLine("Ghi thêm thành công vào file.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi ghi thêm vào file: " + ex.Message);
            }
        }

    }
}

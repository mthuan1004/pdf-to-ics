using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test
{
    internal class LichDayGiangVien
    {
        public GiangVien GiangVien { get; set; }
        public List<MonHoc> LichDay { get; set; }
        public LichDayGiangVien(GiangVien giangVien, List<MonHoc> lichDay)
        {
            GiangVien = giangVien;
            LichDay = lichDay;
        }
        public LichDayGiangVien() { 
        }

       public static List<LichDayGiangVien> ProcessExcelData(string filePath)
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
                        var giangVien = GiangVien.ProcessGiangVien(cellValue);
                        var lichDay = MonHoc.ProcessMonHocData(reader); // Refactored into a new method
                        if(lichDay == null)
                        {
                            return lichDayGiangViens;
                        }
                        else
                        {
                        lichDayGiangViens.Add(new LichDayGiangVien(giangVien, lichDay)); // Thêm thông tin lịch dạy của giảng viên vào danh sách
                        }
                    }
                }
            }

            return lichDayGiangViens;
        }
    }
}

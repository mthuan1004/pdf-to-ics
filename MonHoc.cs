using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    internal class MonHoc
    {
        public string MaMonHoc { get; set; }
        public string Nhom { get; set; }
        public string TenMonHoc { get; set; }
        public string ThoiGian { get; set; }
        public string ThoiGianBD { get; set; }
        public string ThoiGianKT { get; set; }
        public string Phong { get; set; }
        public string TietHoc { get; set; }
        public string Thu { get; set; }
        public string SiSo { get; set; }
        public string Lop { get; set; }


        public MonHoc(string maMonHoc, string nhom, string tenMonHoc, string thoiGian, string thoiGianBD, string thoiGianKT, string phong, string tietHoc, string thu, string siSo, string lop)
        {
            MaMonHoc = maMonHoc;
            Nhom = nhom;
            TenMonHoc = tenMonHoc;
            ThoiGian = thoiGian;
            ThoiGianBD = thoiGianBD;
            ThoiGianKT = thoiGianKT;
            Phong = phong;
            TietHoc = tietHoc;
            Thu = thu;
            SiSo = siSo;
            Lop = lop;

        }
        public MonHoc()
        {


        }

        public override string ToString()
        {
            return $"Mã môn học: {MaMonHoc}" +
                $"\nNhóm: {Nhom}" +
                $"\nTên môn học: {TenMonHoc}" +
                $"\nLớp: {Lop}" +
                $"\nSĩ số: {SiSo}" +
                $"\nThứ: {Thu}" +
                $"\nTiết học: {TietHoc}" +
                $"\nPhòng: {Phong}" +
                $"\nThời gian: {ThoiGian}" +
                $"\nThời gian bắt đầu: {ThoiGianBD}" +
                $"\nThời gian kết thúc: {ThoiGianKT}" + "\n";
        }

        public static MonHoc ProcessMonHoc(string cellValue)
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

        public static List<MonHoc> ProcessMonHocData(IExcelDataReader reader)
        {
            var lichDay = new List<MonHoc>();
            XuLyFile writeFile = new XuLyFile();
            List<string> maMonList = new XuLyFile().ReadCodesFromFile();

            
            while (reader.Read())
            {
                string cellValue = reader.GetString(0);

                if (cellValue.StartsWith("Thời gian học:"))
                {
                    break; // Dừng khi đạt đến dòng thời gian
                }
                else if (CompareMaMon(cellValue, maMonList))
                {
                    var monHoc = ProcessMonHoc(cellValue); // Xử lý môn học
                    lichDay.Add(monHoc); // Thêm vào lịch
                }
                else
                {

                    writeFile.AppendToFile(cellValue); // Ghi vào file
                }

            }
            return lichDay;
        }

        public static bool CompareMaMon(string cell, List<string> maMonList)
        {
            return maMonList.Any(ma => cell.StartsWith(ma));
        }
    }
}
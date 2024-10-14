using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test
{
    internal class GiangVien
    {
        public string MaGiangVien { get; set; }
        public string TenGiangVien { get; set; }
        public string SoDienThoai { get; set; }
        public string Khoa { get; set; }
        public GiangVien(string maGiangVien, string tenGiangVien, string soDienThoai, string khoa)
        {
            MaGiangVien = maGiangVien;
            TenGiangVien = tenGiangVien;
            SoDienThoai = soDienThoai;
            Khoa = khoa;
        }

        public static GiangVien ProcessGiangVien(string cellValue)
        {
            string tenGiangVien = "";
            string maGiangVien = "";
            string[] separators = new string[] { "Cán bộ giảng dạy: ", "CBGD ", " (" };
            string[] result = cellValue.Split(separators, StringSplitOptions.RemoveEmptyEntries);


            // Trích xuất thông tin
            List<string> lstr = result.ToList();
            if (lstr.Count() < 2)
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
    }
}
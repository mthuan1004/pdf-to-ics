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

        
    }
}
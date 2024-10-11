using System;
using System.Collections.Generic;
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
    }
}

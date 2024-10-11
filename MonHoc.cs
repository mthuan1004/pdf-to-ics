using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
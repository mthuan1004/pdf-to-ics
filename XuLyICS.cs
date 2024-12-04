using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    internal class XuLyICS
    {
        public XuLyICS() 
        {
        
        }

        public static void GenerateIcsFile(List<LichDayGiangVien> lichDayGiangViens, string outputDirectory)
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
                        if (monHoc == null)
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
                        writer.WriteLine($"SUMMARY:[{monHoc.Nhom}]{monHoc.TenMonHoc}");
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
        
    }
}

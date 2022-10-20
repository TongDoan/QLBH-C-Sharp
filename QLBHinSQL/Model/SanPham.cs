using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QLBHinSQL.Model
{
    internal class SanPham
    {
         private string maSP;
         private string tenSP;
         private string ngaySX;
         private string ngayHH;
         private string donVi;
         private string donGia;
         private string ghiChu;

        public SanPham(string maSP, string tenSP, string ngaySX, string ngayHH, string donVi, string donGia, string ghiChu)
        {
            this.maSP = maSP;
            this.tenSP = tenSP;
            this.ngaySX = ngaySX;
            this.ngayHH = ngayHH;
            this.donVi = donVi;
            this.donGia = donGia;
            this.ghiChu = ghiChu;
        }

        public string MaSP { get => maSP; set => maSP = value; }
        public string TenSP { get => tenSP; set => tenSP = value; }
        public string NgaySX { get => ngaySX; set => ngaySX = value; }
        public string NgayHH { get => ngayHH; set => ngayHH = value; }
        public string DonVi { get => donVi; set => donVi = value; }
        public string DonGia { get => donGia; set => donGia = value; }
        public string GhiChu { get => ghiChu; set => ghiChu = value; }
    }
}

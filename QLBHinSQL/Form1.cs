using QLBHinSQL.data;
using QLBHinSQL.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;
using Syncfusion;
using System.Collections;
using System.Drawing.Imaging;
using Syncfusion.XlsIO;

namespace QLBHinSQL
{
    public partial class Form1 : Form
    {
        
        Connection db = new Connection();
        public Form1()
        {
            InitializeComponent();
        }
        private SanPham sanPham()
        {
            string maSP = txtMaSP.Text.Trim();
            string tenSP = txtTenSP.Text.Trim();
            string ngaySX = DTPNgaySX.Value.ToString("yyyy-MM-dd");
            string ngayHH = DTPNgayHH.Value.ToString("yyyy-MM-dd");
            string donVi = txtDonVi.Text.Trim();
            string donGia = txtDonGia.Text.Trim();
            string GhiChu = txtGhiChu.Text.Trim();
            return new SanPham(maSP, tenSP, ngaySX, ngayHH, donVi, donGia, GhiChu);
        }
        
        private void stateGB(bool t) 
        {
            gbChiTiet.Enabled = t;
        }
       
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }
        private void StateDisableButton(bool t)
        {
            btnThem.Enabled = t;
            btnSua.Enabled = t;
            btnXoa.Enabled = t;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            lbTieuDe.Text = "THÊM MẶT HÀNG";
            resetForm();
            stateGB(true);
            StateDisableButton(false);
            txtMaSP.Enabled = true;
            btnThem.Enabled = true;
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
           
            if (txtMaSP.Text == "")
            {
                MessageBox.Show("Vui lòng chọn sản phẩm muốn sửa!", "Thông báo");
            }
            else
            {
                lbTieuDe.Text = "CẬP NHẬT MẶ HÀNG";
                stateGB(true);
                StateDisableButton(false);
                btnSua.Enabled = true;
                txtMaSP.Enabled = false;
                stateHD(true);
            }
        }
        private void stateHD(bool t)
        {
            txtTenSP.Enabled = t;
            DTPNgaySX.Enabled = t;
            DTPNgayHH.Enabled = t;
            txtDonVi.Enabled = t;
            txtDonGia.Enabled = t;
            txtGhiChu.Enabled = t;
        }
        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (txtMaSP.Text == "")
            {
                MessageBox.Show("Vui lòng chọn sản phẩm muốn xóa!", "Thông báo");
            }
            else if (MessageBox.Show("Bạn có chắc chắn xóa mã mặt hàng " + txtMaSP.Text + " không? Nếu có ấn nút Lưu, không thì ấn nút Hủy", "Xóa sản phẩm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                lbTieuDe.Text = "XÓA MẶT HÀNG";
                stateGB(true);
                StateDisableButton(false);
                btnXoa.Enabled = true;
                txtMaSP.Enabled = false;
                stateHD(false);
            }
            
        }
        private bool check()
        {
            if(txtMaSP.Text.Trim() == "")
            {
                MessageBox.Show("Mã sản phẩm không được để trống", "Thông báo");
                return false;
            }
            if(txtTenSP.Text.Trim() == "")
            {
                MessageBox.Show("Tên sản phẩm không được để trống", "Thông báo");
                return false;
            } 
            if(DTPNgaySX.Value > DateTime.Now)
            {
                MessageBox.Show("Ngày sản xuất phải lớn hơn ngày hiện tại", "Thông báo");
                return false;
            }
            TimeSpan space = DTPNgayHH.Value.Subtract(DTPNgaySX.Value);
            if(space.TotalDays < 30)
            {
                MessageBox.Show("Ngày hết hạn không được nhỏ hơn 30 ngày", "Thông báo");
                return false;
            }
            if (txtDonVi.Text.Trim() == "")
            {
                MessageBox.Show("Đơn vị không được để trống", "Thông báo");
                return false;
            }
            if (txtDonGia.Text.Trim() == "")
            {
                MessageBox.Show("Đơn giá không được để trống", "Thông báo");
                return false;
            }
            return true;
        }
        private void btnLuu_Click(object sender, EventArgs e)
        {

            if (btnThem.Enabled && check())
            {
                SanPham sp = sanPham();
                string query = $"Insert into QLBH values('{sp.MaSP}',N'{sp.TenSP}','{sp.NgaySX}','{sp.NgayHH}','{sp.DonVi}',{sp.DonGia},N'{sp.GhiChu}')";
                if (MessageBox.Show("Bạn có muốn thêm tên sản phẩm không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes){
                    try
                    {
                        db.Excute(query);
                        MessageBox.Show("Thêm sản phẩm thành công!","Thông báo");
                        Form1_Load(sender, e);
                        resetForm();
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Error: " + "Mã sản phẩm đã tồn tại vui lòng nhập mã khác!","Thông báo");
                    }
                }
            }
            if (btnSua.Enabled && check())
            {
                SanPham sp = sanPham();
                string query = $"Update QLBH set TenSP = N'{sp.TenSP}',NgaySX = '{sp.NgaySX}',NgayHH = '{sp.NgayHH}',DonVi = '{sp.DonVi}',DonGia = '{sp.DonGia}',GhiChu = N'{sp.GhiChu}' where MaSP = '{txtMaSP.Text.Trim()}'";
                try
                {
                    db.Excute(query);
                    MessageBox.Show("Update thành công!", "Thông báo");
                    Form1_Load(sender, e);
                   

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }    
            }
            if (btnXoa.Enabled)
            {
                string query = $"Delete from QLBH where MaSP = N'{txtMaSP.Text.Trim()}'";
                try
                {
                    db.Excute(query);
                    MessageBox.Show("Xóa thành công!", "Thông báo");
                    Form1_Load(sender, e);
                    resetForm();
                    stateGB(false);
                    StateDisableButton(true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            lbTieuDe.Text = "QUẢN LÝ SẢN PHẨM";
            stateGB(false);
            StateDisableButton(true);
        }

        private void txtDonGia_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dgvKetQua.DataSource = db.table("Select * from QLBH");
        }
      
        private void resetForm()
        {
            txtMaSP.Text = "";
            txtTenSP.Text = "";
            DTPNgaySX.Value = DateTime.Now;
            DTPNgayHH.Value = DateTime.Now;
            txtDonVi.Text = "";
            txtDonGia.Text = "";
            txtGhiChu.Text = "";
        }
        
        private void btnTimkiem_Click(object sender, EventArgs e)
        {
            lbTieuDe.Text = "TÌM KIẾM MẶT HÀNG";
            if (txtTKTenSP.Text == "")
            {
                MessageBox.Show("Vui lòng nhập tên sản phẩm muốn tìm!", "Thông báo");
                Form1_Load(sender, e);
            }
            else
            {
                string query = $"Select * from QLBH where TenSP = N'{txtTKTenSP.Text.Trim()}'";
                try
                {
                    dgvKetQua.DataSource = db.table(query);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void dgvKetQua_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            stateGB(false);
            StateDisableButton(false);
            btnSua.Enabled = true;
            btnXoa.Enabled = true;
            txtMaSP.Text = dgvKetQua.SelectedRows[0].Cells[0].Value.ToString();
            txtTenSP.Text = dgvKetQua.SelectedRows[0].Cells[1].Value.ToString();
            DTPNgaySX.Value = DateTime.Parse(dgvKetQua.SelectedRows[0].Cells[2].Value.ToString());
            DTPNgayHH.Value = DateTime.Parse(dgvKetQua.SelectedRows[0].Cells[3].Value.ToString());
            txtDonVi.Text = dgvKetQua.SelectedRows[0].Cells[4].Value.ToString();
            txtDonGia.Text = dgvKetQua.SelectedRows[0].Cells[5].Value.ToString();
            txtGhiChu.Text = dgvKetQua.SelectedRows[0].Cells[6].Value.ToString();
        }

        private void Export(DataTable dataTable, string path)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                worksheet.ImportDataTable(dataTable, true, 1, 1);

                worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs(path);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel";
            saveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Excel 2003(*.xls)|*.xls";
            if(saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DataTable dataTable = db.table("Select * from QLBH");
                    Export(dataTable,saveFileDialog.FileName);
                    MessageBox.Show("Xuất file thành công!","Thông báo");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo", MessageBoxButtons.YesNo,
               MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                        this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 fom = new Form2();
            fom.ShowDialog();
        }
    }
}

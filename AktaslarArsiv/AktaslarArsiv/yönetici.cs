using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace AktaslarArsiv
{
    public partial class yönetici : Form
    {

        String id;
        public yönetici(String _id)
        {
            InitializeComponent();

            id = _id;
        }
        OleDbConnection baglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=aktasDB.accdb");
        OleDbDataAdapter da;
        DataTable dt;
        string sql = "SELECT * FROM aktasArsiv";

        void Listele()
        {
            da = new OleDbDataAdapter(sql, baglanti);
            dt = new DataTable();
            baglanti.Open();
            da.Fill(dt);
            baglanti.Close();
            dataGrid.DataSource = dt;
        }

        bool dragging;

        Point offset;
        private void label5_MouseMove(object sender, MouseEventArgs e)
        {
            label5.BackColor = Color.Tomato;
        }

        private void label5_MouseLeave(object sender, EventArgs e)
        {
            label5.BackColor = Color.Transparent;
        }

        private void yönetici_MouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;
        }

        private void yönetici_MouseDown(object sender, MouseEventArgs e)
        {
            dragging = true;
            offset = e.Location;
        }

        private void yönetici_MouseMove(object sender, MouseEventArgs e)
        {
            if (dragging)
            {
                Point currentScreenPos = PointToScreen(e.Location);
                Location = new
                Point(currentScreenPos.X - offset.X,
                currentScreenPos.Y - offset.Y);
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void yönetici_Load(object sender, EventArgs e)
        {
            k_adi.Text = id;
            comboBox1.SelectedItem = comboBox1.Items[0];
            Listele();


        }

        private void save_btn_Click(object sender, EventArgs e)
        {
            int esas_id;
            Random random = new Random();

            esas_id = random.Next(0, 10000000);

            try
            {

                save(esas_id);
            }
            catch
            {
                esas_id = random.Next(0, 10000000);
                save(esas_id);

            }

        }

        public void save(int esas_id)
        {
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("INSERT INTO aktasArsiv (id,MusteriAdi,Ada,Parsel,MahalleKoy,IsTanim,Tarih,Fiyat,TelefonNo,Aciklama,islemiYapan) VALUES (@id,@MusteriAdi,@Ada,@Parsel,@MahalleKoy,@IsTanim,@Tarih,@Fiyat,@TelefonNo,@Aciklama,@islemiYapan)", baglanti);

            komut.Parameters.AddWithValue("@id", esas_id.ToString());
            komut.Parameters.AddWithValue("@MusteriAdi", txt_MusteriAdi.Text);
            komut.Parameters.AddWithValue("@Ada", txt_Ada.Text);
            komut.Parameters.AddWithValue("@Parsel", txt_Parsel.Text);
            komut.Parameters.AddWithValue("@MahalleKoy", txt_MahalleKoy.Text);
            komut.Parameters.AddWithValue("@IsTanim", txt_IsTanim.Text);
            komut.Parameters.AddWithValue("@Tarih", txt_Tarih.Text);
            komut.Parameters.AddWithValue("@Fiyat", txt_Fiyat.Text);
            komut.Parameters.AddWithValue("@TelefonNo", txt_TelefonNo.Text);
            komut.Parameters.AddWithValue("@Aciklama", txt_Aciklama.Text);
            komut.Parameters.AddWithValue("@islemiYapan", id);//işlemi yapan id alıyoruz



            komut.ExecuteNonQuery();
            baglanti.Close();
            txt_MusteriAdi.Clear();
            txt_Ada.Clear();
            txt_Parsel.Clear();
            txt_MahalleKoy.Clear();
            txt_IsTanim.Clear();
            txt_Tarih.Clear();
            txt_Fiyat.Clear();
            txt_TelefonNo.Clear();
            txt_Aciklama.Clear();

            Listele();


        }


        private void dataGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            txt_id.Text = dataGrid.CurrentRow.Cells[0].Value.ToString();
            txt_MusteriAdi.Text = dataGrid.CurrentRow.Cells[1].Value.ToString();
            txt_Ada.Text = dataGrid.CurrentRow.Cells[2].Value.ToString();
            txt_Parsel.Text = dataGrid.CurrentRow.Cells[3].Value.ToString();
            txt_MahalleKoy.Text = dataGrid.CurrentRow.Cells[4].Value.ToString();
            txt_IsTanim.Text = dataGrid.CurrentRow.Cells[5].Value.ToString();
            txt_Tarih.Text = dataGrid.CurrentRow.Cells[6].Value.ToString();
            txt_Fiyat.Text = dataGrid.CurrentRow.Cells[7].Value.ToString();
            txt_TelefonNo.Text = dataGrid.CurrentRow.Cells[8].Value.ToString();
            txt_Aciklama.Text = dataGrid.CurrentRow.Cells[9].Value.ToString();
        }

        private void btn_Sil_Click(object sender, EventArgs e)
        {

            string message = "Silmek istediğinizden emin misiniz ?";
            string title = "Sil";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                OleDbCommand komut = new OleDbCommand("DELETE FROM aktasArsiv WHERE id = @id", baglanti);
                komut.Parameters.AddWithValue("@id", txt_id.Text);
                baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();

                Listele();
            }




        }

        private void btn_Guncelle_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("UPDATE aktasArsiv SET MusteriAdi='" + txt_MusteriAdi.Text + "',Ada='" + txt_Ada.Text + "',Parsel='" + txt_Parsel.Text + "',MahalleKoy='" + txt_MahalleKoy.Text + "',IsTanim='" + txt_IsTanim.Text + "',Tarih='" + txt_Tarih.Text + "',Fiyat='" + txt_Fiyat.Text + "',TelefonNo='" + txt_TelefonNo.Text + "',Aciklama='" + txt_Aciklama.Text + "' WHERE id='" + txt_id.Text + "'", baglanti);

            komut.ExecuteNonQuery();
            baglanti.Close();
            Listele();



        }

        private void ara_text_TextChanged(object sender, EventArgs e)
        {

            try
            {
                DataView dv = dt.DefaultView;
                string veri = string.Format("{0} LIKE '" + ara_text.Text + "%'", comboBox1.Text);
                dv.RowFilter = veri;
            }
            catch
            {

            }
        }

        private void excl_Btn_Click(object sender, EventArgs e)
        {


            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "veriler.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                
                ToCsV(dataGrid, sfd.FileName); 
            }

        }
        private void ToCsV(DataGridView dGV, string filename)
        {
            string stOutput = "";
            
            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            
            for (int i = 0; i < dGV.RowCount; i++)
            {
                string stLine = "";
                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); 
            bw.Flush();
            bw.Close();
            fs.Close();
        }
    }
}


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

using Excel = Microsoft.Office.Interop.Excel;

using Microsoft.Office.Interop.Excel;

namespace BÜT_Mehmed_Akif_Aydemir
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            this.button6.BackColor = System.Drawing.Color.Green;

            this.button7.BackColor = System.Drawing.Color.Green;

            this.button8.BackColor = System.Drawing.Color.Green;

            this.button9.BackColor = System.Drawing.Color.Green;

            this.button10.BackColor = System.Drawing.Color.Green;

            this.button11.BackColor = System.Drawing.Color.Green;

            this.button12.BackColor = System.Drawing.Color.Green;
            this.button13.BackColor = System.Drawing.Color.Green;
            this.button14.BackColor = System.Drawing.Color.Green;
            this.button15.BackColor = System.Drawing.Color.Green;
            this.button16.BackColor = System.Drawing.Color.Green;
            this.button17.BackColor = System.Drawing.Color.Green;
            this.button18.BackColor = System.Drawing.Color.Green;
            this.button19.BackColor = System.Drawing.Color.Green;
            this.button20.BackColor = System.Drawing.Color.Green;
            this.button21.BackColor = System.Drawing.Color.Green;
            this.button4.BackColor = System.Drawing.Color.Green;
            this.button5.BackColor = System.Drawing.Color.Green;




            dataGridView1.Columns.Add(label2.Text, label2.Text);
            dataGridView1.Columns.Add(label3.Text, label3.Text);
            dataGridView1.Columns.Add(label4.Text, label4.Text);
            dataGridView1.Columns.Add(label5.Text, label5.Text);
            dataGridView1.Columns.Add(label6.Text, label6.Text);
            dataGridView1.Columns.Add(label7.Text, label7.Text);
            dataGridView1.Columns.Add(label8.Text, label8.Text);
            dataGridView1.Columns.Add(label9.Text, label9.Text);
            dataGridView1.Columns.Add(label10.Text, label10.Text);
            dataGridView1.Columns.Add(label11.Text, label11.Text);
            dataGridView1.Columns.Add(label12.Text, label12.Text);
            dataGridView1.Columns.Add(label13.Text, label13.Text);
            dataGridView1.Columns.Add(label18.Text, label18.Text);
            dataGridView1.Columns.Add(label19.Text, label19.Text);
            dataGridView1.Columns.Add(label20.Text, label20.Text);
        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://www.thy.com");
        }






        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        private string TcDogrulama(string tcNo)
        {
            string durum = "";
            try
            {
                if (tcNo != "")
                {
                    if (tcNo.Length == 11)
                    {
                        char[] rakamlar = tcNo.ToCharArray();
                        int kural1 = 0, hane11 = rakamlar[10], hane10 = rakamlar[9];

                        for (int i = 0; i < 10; i++)
                        {
                            kural1 += Convert.ToInt32(rakamlar[i].ToString());
                        }
                        char[] birlerbasamagikural1 = kural1.ToString().ToCharArray();

                        int kural2tek = 0, kural2cift = 0;

                        for (int i = 0; i < 10; i += 2)
                        {
                            kural2tek += Convert.ToInt32(rakamlar[i].ToString());
                        }
                        for (int i = 1; i < 9; i += 2)
                        {
                            kural2cift += Convert.ToInt32(rakamlar[i].ToString());
                        }
                        char[] birlerbasamagikural2 = ((7 * kural2tek) + (9 * kural2cift)).ToString().ToCharArray();

                        int kural3 = 0;

                        for (int i = 0; i < 10; i += 2)
                        {
                            kural3 += Convert.ToInt32(rakamlar[i].ToString());
                        }
                        char[] birlerbasamagikural3 = (8 * kural3).ToString().ToCharArray();

                        if ((birlerbasamagikural1[birlerbasamagikural1.Length - 1] == hane11) && (birlerbasamagikural2[birlerbasamagikural2.Length - 1] == hane10) && (birlerbasamagikural3[birlerbasamagikural3.Length - 1] == hane11))
                        {
                            durum = "Kimlik Numarası Geçerli";
                        }
                        else
                        {
                            durum = "Kimlik Numarası Geçerli Değildir";
                        }

                        textBox1.Focus();
                    }
                    else
                    {
                        durum = "TC Kimlik Numaranızı Eksik Girdiniz \n" +
                            " Lütfen Kontrol Ediniz!!!";
                    }
                }
                else
                {
                    durum = "Lütfen TC Kimlik Numaranızı Giriniz!!!";
                }
            }
            catch (Exception ex)
            {
                durum = ex.Message;
            }
            return durum;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            

        }




        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://www.rixos.com");
        }

        private void HESAPLA_Click(object sender, EventArgs e)
        {
            float fiyat = 500;

            int adet;

            float tutar;

            fiyat = Convert.ToSingle(fiyat);

            adet = Convert.ToInt32(numericUpDown1.Text);

            tutar = adet * fiyat;

            textBox5.Text = tutar.ToString() + "₺";
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }




        
        private void button3_Click_1(object sender, EventArgs e)
        {
            OleDbCommand komut = new OleDbCommand();



            Excel.Application excel = new Excel.Application();

            excel.Visible = true;

            object Missing = Type.Missing;

            Workbook workbook = excel.Workbooks.Add(Missing);

            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            int StartCol = 1;

            int StartRow = 1;

            for (int j = 0; j < dataGridView1.Columns.Count; j++)

            {

                Microsoft.Office.Interop.Excel.Range myRange =

                    (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];

                myRange.Value2 = dataGridView1.Columns[j].HeaderText;

            }

            StartRow++;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)

            {

                for (int j = 0; j < dataGridView1.Columns.Count; j++)

                {

                    try

                    {

                        Microsoft.Office.Interop.Excel.Range myRange =

                            (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];

                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;

                        myRange.Select();

                    }

                    catch

                    {

                        ;

                    }

                }

            }



        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)

        {



        }

        private void button22_Click(object sender, EventArgs e)
        {

            dataGridView1.Rows.Add(textBox1.Text, textBox2.Text, comboBox2.Text, comboBox3.Text, textBox3.Text, textBox4.Text, comboBox4.Text, dateTimePicker1.Text, dateTimePicker2.Text, numericUpDown1.Text, comboBox5.Text, textBox5.Text, numericUpDown2.Text, comboBox1.Text, textBox9.Text);


        }
        



        
        private void button6_Click_1(object sender, EventArgs e)
        {

            this.button6.BackColor = System.Drawing.Color.Red;

            button6.Enabled = false;

            textBox9.Text = textBox9.Text + button6.Text;
        }

        private void button7_Click(object sender, EventArgs e)
        {

            this.button7.BackColor = System.Drawing.Color.Red;

            button7.Enabled = false;

            textBox9.Text = textBox9.Text + button7.Text;
        }

        private void button8_Click(object sender, EventArgs e)
        {

            this.button8.BackColor = System.Drawing.Color.Red;

            button8.Enabled = false;

            textBox9.Text = textBox9.Text + button8.Text;
        }

        private void button9_Click(object sender, EventArgs e)
        {

            this.button9.BackColor = System.Drawing.Color.Red;

            button9.Enabled = false;

            textBox9.Text = textBox9.Text + button9.Text;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.button10.BackColor = System.Drawing.Color.Red;

            button8.Enabled = false;

            textBox9.Text = textBox9.Text + button10.Text;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            this.button11.BackColor = System.Drawing.Color.Red;

            button11.Enabled = false;

            textBox9.Text = textBox9.Text + button11.Text;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            this.button12.BackColor = System.Drawing.Color.Red;

            button12.Enabled = false;

            textBox9.Text = textBox9.Text + button12.Text;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            this.button13.BackColor = System.Drawing.Color.Red;

            button13.Enabled = false;

            textBox9.Text = textBox9.Text + button13.Text;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            this.button14.BackColor = System.Drawing.Color.Red;

            button14.Enabled = false;

            textBox9.Text = textBox9.Text + button14.Text;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            this.button15.BackColor = System.Drawing.Color.Red;

            button15.Enabled = false;

            textBox9.Text = textBox9.Text + button15.Text;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            this.button16.BackColor = System.Drawing.Color.Red;

            button16.Enabled = false;

            textBox9.Text = textBox9.Text + button16.Text;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.button17.BackColor = System.Drawing.Color.Red;

            button17.Enabled = false;

            textBox9.Text = textBox9.Text + button17.Text;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            this.button18.BackColor = System.Drawing.Color.Red;

            button18.Enabled = false;

            textBox9.Text = textBox9.Text + button18.Text;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            this.button19.BackColor = System.Drawing.Color.Red;

            button19.Enabled = false;

            textBox9.Text = textBox9.Text + button19.Text;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            this.button20.BackColor = System.Drawing.Color.Red;

            button20.Enabled = false;

            textBox9.Text = textBox9.Text + button20.Text;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            this.button21.BackColor = System.Drawing.Color.Red;

            button21.Enabled = false;

            textBox9.Text = textBox9.Text + button21.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.button4.BackColor = System.Drawing.Color.Red;

            button4.Enabled = false;

            textBox9.Text = textBox9.Text + button4.Text;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.button5.BackColor = System.Drawing.Color.Red;

            button5.Enabled = false;

            textBox9.Text = textBox9.Text + button5.Text;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            label24.Text = (TcDogrulama(textBox1.Text));
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

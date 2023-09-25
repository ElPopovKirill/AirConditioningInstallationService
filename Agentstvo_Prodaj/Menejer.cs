using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Agentstvo_Prodaj
{
    public partial class Menejer : Form
    {
        public Menejer()
        {
            InitializeComponent();
        }

        private void Menejer_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet11.Сотрудники". При необходимости она может быть перемещена или удалена.
            this.сотрудникиTableAdapter1.Fill(this.kyrsachDataSet11.Сотрудники);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet10.Услуги". При необходимости она может быть перемещена или удалена.
            this.услугиTableAdapter1.Fill(this.kyrsachDataSet10.Услуги);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet9.Способы_оплаты". При необходимости она может быть перемещена или удалена.
            this.способы_оплатыTableAdapter1.Fill(this.kyrsachDataSet9.Способы_оплаты);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet8.Кондиционеры". При необходимости она может быть перемещена или удалена.
            this.кондиционерыTableAdapter3.Fill(this.kyrsachDataSet8.Кондиционеры);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet7.Клиенты". При необходимости она может быть перемещена или удалена.
            this.клиентыTableAdapter1.Fill(this.kyrsachDataSet7.Клиенты);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet6.Кондиционеры". При необходимости она может быть перемещена или удалена.
            this.кондиционерыTableAdapter2.Fill(this.kyrsachDataSet6.Кондиционеры);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet5.Кондиционеры". При необходимости она может быть перемещена или удалена.
            this.кондиционерыTableAdapter1.Fill(this.kyrsachDataSet5.Кондиционеры);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet4.Услуги". При необходимости она может быть перемещена или удалена.
            this.услугиTableAdapter.Fill(this.kyrsachDataSet4.Услуги);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet3.Способы_оплаты". При необходимости она может быть перемещена или удалена.
            this.способы_оплатыTableAdapter.Fill(this.kyrsachDataSet3.Способы_оплаты);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet2.Сотрудники". При необходимости она может быть перемещена или удалена.
            this.сотрудникиTableAdapter.Fill(this.kyrsachDataSet2.Сотрудники);          
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kyrsachDataSet.Клиенты". При необходимости она может быть перемещена или удалена.
            this.клиентыTableAdapter.Fill(this.kyrsachDataSet.Клиенты);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            panel1.Visible = false;
            panel2.Visible = false;
            panel4.Visible = false;
            panel6.Visible = false;
            panel9.Visible = false;

            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            String query = "SELECT * from Клиенты";
            str.Open();
            SqlDataAdapter a = new SqlDataAdapter(query, str);
            DataSet ds = new DataSet();
            a.Fill(ds);
            dataGridView3.DataSource = ds.Tables[0];
            str.Close();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            String query = "select ID_Договора, Сотрудники.ФИО as сотрудник,Клиенты.ФИО as клиент,Кондиционеры.Название as кондиционер,Способы_оплаты.Способ_оплтаы as способ_оплаты,Услуги.Услуга as услуга,Адрес_Получателя,Количество_кондицеонеров,Договора.Цена,Дата_заключения from Договора join Сотрудники on Сотрудники.ID_сотрудника = Договора.ID_Сотрудника join Клиенты on Клиенты.ID_Клиента = Договора.ID_Клиента join Кондиционеры on Кондиционеры.ID_Кондиционера = Договора.ID_Кондиционера join Способы_оплаты on Способы_оплаты.ID_Способа_оплаты = Договора.ID_Способа_оплаты join Услуги on Услуги.ID_Услуги = Договора.ID_Услуги";
            str.Open();
            SqlDataAdapter a = new SqlDataAdapter(query, str);
            DataSet ds = new DataSet();
            a.Fill(ds);
            dataGridView2.DataSource = ds.Tables[0];
            str.Close();
            panel3.Visible = false;
            panel1.Visible = false;
            panel2.Visible = true;
            panel4.Visible = false;
            panel6.Visible = false;
            panel9.Visible = false;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel3.Visible = false;        
            panel2.Visible = false;
            panel4.Visible = false;
            panel6.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel3.Visible = false;
            panel2.Visible = false;
            panel4.Visible = true;
            panel6.Visible = false;
            panel3.Visible = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel3.Visible = true;
            button3_Click(sender, e);
        }

        private void Menejer_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();
            Application.Exit();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = true;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            panel5.Visible = false;
            panel1.Visible = true;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            panel5.Visible = true;
            panel1.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                int id = Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value);
                SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
                str.Open();
                SqlCommand command = new SqlCommand($"DELETE Клиенты WHERE (ID_Клиента) = " + id + "", str);
                command.ExecuteNonQuery();
                MessageBox.Show("Клиент успешно удалён!");
                str.Close();
                button3_Click(sender, e);
            }
            catch (Exception)
            {
                MessageBox.Show("У клиента активный договор!");
               
            }
           
        }

        private void button16_Click(object sender, EventArgs e)
        {
            panel6.Visible = false;
            panel3.Visible = true;
        }


      
        string FIO = "";
        string Date_birthday = "";
        string Number = "";
        string passport = "";
        private void button7_Click(object sender, EventArgs e)
        {
            panel6.Visible = true;
            panel3.Visible = false;
            int id = Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value);
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str.Open();
            SqlCommand command = new SqlCommand("select ID_Клиента,ФИО,Серия_номер_паспорта,Дата_рождения,Телефон from Клиенты WHERE ID_Клиента = " + id + "", str);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                
                FIO = reader[1].ToString();
                Date_birthday = reader[3].ToString();
                Number = reader[4].ToString();
                passport = reader[2].ToString();
              

            }
            reader.Close();
            str.Close();
           
            textBox8.Text = FIO;
            dateTimePicker2.Text = Date_birthday;
            textBox7.Text = Number;         
            textBox6.Text = passport;


        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (textBox4.Text.Length > 0 && textBox3.Text.Length > 0 && textBox2.Text.Length > 0)
            {
               

                SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
                str.Open();
                SqlCommand command = new SqlCommand($"insert into  Клиенты(ФИО,Серия_номер_паспорта,Дата_Рождения,Телефон) values('" + textBox4.Text + "', '" + textBox3.Text + "', '" + dateTimePicker1.Value + "', '" + textBox2.Text + "');", str);
                command.ExecuteNonQuery();
                MessageBox.Show("Клиент успешно добавлен!");
                str.Close();
                textBox4.Text = "";
                textBox3.Text = "";
                textBox2.Text = "";
                panel4.Visible = false;
            }
            else
            {
                MessageBox.Show("Заполните все данные!");
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value);
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str.Open();
            SqlCommand command = new SqlCommand($"update Клиенты set ФИО='" + textBox8.Text + "' , Серия_номер_паспорта = '" + textBox6.Text + "',Дата_рождения='" + dateTimePicker2.Value + "',Телефон='" + textBox7.Text + "' WHERE ID_Клиента = " + id + "", str);
            command.ExecuteNonQuery();
            MessageBox.Show("Информация успешно изменена!");
            str.Close();     
            panel6.Visible = false;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            panel7.Visible = false;
            panel1.Visible = true;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            panel7.Visible = true;
            panel1.Visible = false;

            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            String query = "SELECT * from Услуги";
            str.Open();
            SqlDataAdapter a = new SqlDataAdapter(query, str);
            DataSet ds = new DataSet();
            a.Fill(ds);
            dataGridView4.DataSource = ds.Tables[0];
            str.Close();



        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                int id = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value);
                SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
                str.Open();
                SqlCommand command = new SqlCommand($"DELETE Договора WHERE (ID_Договора) = " + id + "", str);
                command.ExecuteNonQuery();
                MessageBox.Show("Договор успешно удалён!");
                str.Close();
                button2_Click(sender, e);
            }
            catch (Exception)
            {
                MessageBox.Show("ошибка!");
              
            }
          
        }

      

        private void panel5_Paint(object sender, PaintEventArgs e)
        {
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            String query = "SELECT * from Кондиционеры";
            str.Open();
            SqlDataAdapter a = new SqlDataAdapter(query, str);
            DataSet ds = new DataSet();
            a.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            str.Close();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {
           
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            string selectchar = textBox9.Text;
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            String query = "select ID_Кондиционера,Название,Размер,Мощность,Цена from Кондиционеры where Название like '%" + selectchar + "%'";
            str.Open();
            SqlDataAdapter a = new SqlDataAdapter(query, str);
            DataSet ds = new DataSet();
            a.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            str.Close();
        }
        string klient;
        string sotrudnik;
        string kondicioner;
        string yslyga;
        int count;
        string vid_oplati;
        string adress;
        string priceq;
        string dateq;
        private void button4_Click(object sender, EventArgs e)
        {
            panel8.Visible = true;
            panel2.Visible = false;

            int id = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value);
            SqlConnection str1 = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str1.Open();
            SqlCommand command1 = new SqlCommand("select  Сотрудники.ФИО as сотрудник,Клиенты.ФИО as клиент,Кондиционеры.Название as кондиционер,Способы_оплаты.Способ_оплтаы as способ_оплаты,Услуги.Услуга as услуга,Адрес_Получателя,Количество_кондицеонеров,Договора.Цена,Дата_заключения from Договора join Сотрудники on Сотрудники.ID_сотрудника = Договора.ID_Сотрудника join Клиенты on Клиенты.ID_Клиента = Договора.ID_Клиента join Кондиционеры on Кондиционеры.ID_Кондиционера = Договора.ID_Кондиционера join Способы_оплаты on Способы_оплаты.ID_Способа_оплаты = Договора.ID_Способа_оплаты join Услуги on Услуги.ID_Услуги = Договора.ID_Услуги where ID_Договора = " + id + "", str1);
            SqlDataReader reader1 = command1.ExecuteReader();
            while (reader1.Read())
            {
                klient = reader1[1].ToString();
                sotrudnik = reader1[0].ToString();
                kondicioner = reader1[2].ToString();
                vid_oplati = reader1[3].ToString();
                yslyga = reader1[4].ToString();
                adress = reader1[5].ToString();
                count = Convert.ToInt32(reader1[6]);
                priceq = reader1[7].ToString();
                dateq = reader1[8].ToString();

            }
            reader1.Close();
            str1.Close();


            comboBox8.Text = klient;
            comboBox10.Text = sotrudnik;
            comboBox7.Text = kondicioner;
            comboBox6.Text = vid_oplati;
            textBox11.Text = count.ToString();
            comboBox2.Text = yslyga;
            textBox10.Text = adress;

        }

        private void button22_Click(object sender, EventArgs e)
        {
            panel8.Visible = false;
            panel2.Visible = true;
        }

        string price = ""; 
        private void button11_Click(object sender, EventArgs e)
        {

            int id = Convert.ToInt32(comboBox5.SelectedValue);
            SqlConnection str1 = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str1.Open();
            SqlCommand command1 = new SqlCommand("select Услуги.Цена from Услуги WHERE ID_Услуги = " + id + "", str1);
            SqlDataReader reader1 = command1.ExecuteReader();
            while (reader1.Read())
            {
                price = Convert.ToString( (Convert.ToInt32(reader1[0].ToString()) * Convert.ToInt32(textBox1.Text)));

            }
            reader1.Close();
            str1.Close();


            if (comboBox1.Text != "" && comboBox1.Text != "" && comboBox4.Text != "" && comboBox5.Text != "")
            {
               
                SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
                str.Open();
                SqlCommand command = new SqlCommand($"insert into Договора(ID_Сотрудника ,ID_Клиента,ID_Кондиционера,ID_Способа_оплаты,ID_Услуги,Адрес_Получателя,Количество_кондицеонеров,Цена) values('" + comboBox9.SelectedValue + "', '" + comboBox1.SelectedValue + "', '" + comboBox3.SelectedValue + "', '" + comboBox4.SelectedValue + "', '" + comboBox5.SelectedValue + "', '" + textBox5.Text + "', '" + Convert.ToInt32(textBox1.Text) + "', '" + price + "');", str);
                command.ExecuteNonQuery();
                MessageBox.Show("Договор успешно добавлен!");
                str.Close();


                button2_Click(sender, e);
            }
            else
            {
                MessageBox.Show("Заполните данные");
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            panel5.Visible = true;
            panel8.Visible = false;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            panel7.Visible = true;
            panel8.Visible = false;
        }

        private void button23_Click(object sender, EventArgs e)
        {

            string redactprice = "";
            int idq = Convert.ToInt32(comboBox2.SelectedValue);
            SqlConnection str1 = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str1.Open();
            SqlCommand command1 = new SqlCommand("select Услуги.Цена from Услуги WHERE ID_Услуги = " + idq + "", str1);
            SqlDataReader reader1 = command1.ExecuteReader();
            while (reader1.Read())
            {
                redactprice = Convert.ToString((Convert.ToInt32(reader1[0].ToString()) * Convert.ToInt32(textBox11.Text)));

            }
            reader1.Close();
            str1.Close();

            int id = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value);
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str.Open();
            SqlCommand command = new SqlCommand($"update Договора set ID_Сотрудника='" + comboBox10.SelectedValue + "' , ID_Клиента = '" + comboBox8.SelectedValue + "',ID_Кондиционера='" + comboBox7.SelectedValue + "',ID_Способа_оплаты='" + comboBox6.SelectedValue + "',ID_Услуги='" + comboBox2.SelectedValue + "',Адрес_Получателя ='" + textBox10.Text + "',Количество_кондицеонеров='" + textBox11.Text + "',Цена='" + redactprice  + "' WHERE ID_Договора = " + id + "", str);
            command.ExecuteNonQuery();
            MessageBox.Show("Информация успешно изменена!");
            str.Close();
            panel8.Visible = false;
            panel2.Visible = true;
            button2_Click(sender, e);

        }

        private void button26_Click(object sender, EventArgs e)
        {
            panel9.Visible = true;
            panel3.Visible = false;
            panel1.Visible = false;
            panel2.Visible = false;
            panel4.Visible = false;
            panel6.Visible = false;
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            String query = "select ID_Договора, Сотрудники.ФИО as сотрудник,Клиенты.ФИО as клиент,Кондиционеры.Название as кондиционер,Способы_оплаты.Способ_оплтаы as способ_оплаты,Услуги.Услуга as услуга,Адрес_Получателя,Количество_кондицеонеров,Договора.Цена,Дата_заключения from Договора join Сотрудники on Сотрудники.ID_сотрудника = Договора.ID_Сотрудника join Клиенты on Клиенты.ID_Клиента = Договора.ID_Клиента join Кондиционеры on Кондиционеры.ID_Кондиционера = Договора.ID_Кондиционера join Способы_оплаты on Способы_оплаты.ID_Способа_оплаты = Договора.ID_Способа_оплаты join Услуги on Услуги.ID_Услуги = Договора.ID_Услуги";
            str.Open();
            SqlDataAdapter a = new SqlDataAdapter(query, str);
            DataSet ds = new DataSet();
            a.Fill(ds);
            dataGridView5.DataSource = ds.Tables[0];
            str.Close();

        }

        private void button24_Click(object sender, EventArgs e)
        {
            panel9.Visible = false; 
        }

        private void button25_Click(object sender, EventArgs e)
        {

            if (textBox14.Text.Length > 0)
            {
              
                SqlConnection str1 = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
                str1.Open();
                SqlCommand command1 = new SqlCommand("select  Сотрудники.ФИО as сотрудник,Клиенты.ФИО as клиент,Кондиционеры.Название as кондиционер,Способы_оплаты.Способ_оплтаы as способ_оплаты,Услуги.Услуга as услуга,Адрес_Получателя,Количество_кондицеонеров,Договора.Цена,Дата_заключения from Договора join Сотрудники on Сотрудники.ID_сотрудника = Договора.ID_Сотрудника join Клиенты on Клиенты.ID_Клиента = Договора.ID_Клиента join Кондиционеры on Кондиционеры.ID_Кондиционера = Договора.ID_Кондиционера join Способы_оплаты on Способы_оплаты.ID_Способа_оплаты = Договора.ID_Способа_оплаты join Услуги on Услуги.ID_Услуги = Договора.ID_Услуги", str1);
                SqlDataReader reader1 = command1.ExecuteReader();
                while (reader1.Read())
                {
                    klient = reader1[1].ToString();
                    sotrudnik = reader1[0].ToString();
                    kondicioner = reader1[2].ToString();
                    vid_oplati = reader1[3].ToString();
                    yslyga = reader1[4].ToString();
                    adress = reader1[5].ToString();
                    count = Convert.ToInt32(reader1[6]);
                    priceq = reader1[7].ToString();
                    dateq = reader1[8].ToString();

                }
                reader1.Close();
                str1.Close();
                // Создаём шаблон документа WORD
                string name_doc = textBox14.Text + ".docx";
                string fileName = @"../Debug/" + name_doc + "";
                var doc = DocX.Create(fileName);
                doc.InsertParagraph("\t\t\t\t\tДоговор оказания услуг\n\n\n");
                //Create Table with 2 rows and 4 columns.  
                Table t = doc.AddTable(9, 2);
                t.Rows[0].Cells[0].Paragraphs.First().Append("Клиент");
                t.Rows[1].Cells[0].Paragraphs.First().Append("Сотрудник");
                t.Rows[2].Cells[0].Paragraphs.First().Append("Кондиционер");
                t.Rows[3].Cells[0].Paragraphs.First().Append("Услуга");
                t.Rows[4].Cells[0].Paragraphs.First().Append("Количество кондиционеров");
                t.Rows[5].Cells[0].Paragraphs.First().Append("Вид оплаты");

                t.Rows[6].Cells[0].Paragraphs.First().Append("Адрес клиента");
                t.Rows[7].Cells[0].Paragraphs.First().Append("Цена");
                t.Rows[8].Cells[0].Paragraphs.First().Append("Дата");


                t.Rows[0].Cells[1].Paragraphs.First().Append(klient);
                t.Rows[1].Cells[1].Paragraphs.First().Append(sotrudnik);
                t.Rows[2].Cells[1].Paragraphs.First().Append(kondicioner);
                t.Rows[3].Cells[1].Paragraphs.First().Append(yslyga);
                t.Rows[4].Cells[1].Paragraphs.First().Append(count.ToString());
                t.Rows[5].Cells[1].Paragraphs.First().Append(vid_oplati);

                t.Rows[6].Cells[1].Paragraphs.First().Append(adress);
                t.Rows[7].Cells[1].Paragraphs.First().Append(priceq);
                t.Rows[8].Cells[1].Paragraphs.First().Append(dateq);
                doc.InsertTable(t);             
                doc.Save();
                Process.Start("WINWORD.EXE", fileName);
            }
            else
            {
                MessageBox.Show("Введите название документа!");
            }
        }

        private void button20_Click_1(object sender, EventArgs e)
        {

        }
    }
}

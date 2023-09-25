using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Agentstvo_Prodaj
{
    public partial class Direktor : Form
    {
        public Direktor()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            String query = "select Сотрудники.ФИО as сотрудник,Клиенты.ФИО as клиент,Кондиционеры.Название as кондицеонер,Способы_оплаты.Способ_оплтаы as способ_оплаты,Услуги.Услуга as услуга,Адрес_Получателя,Количество_кондицеонеров,Договора.Цена,Дата_заключения from Договора join Сотрудники on Сотрудники.ID_сотрудника = Договора.ID_Сотрудника join Клиенты on Клиенты.ID_Клиента = Договора.ID_Клиента join Кондиционеры on Кондиционеры.ID_Кондиционера = Договора.ID_Кондиционера join Способы_оплаты on Способы_оплаты.ID_Способа_оплаты = Договора.ID_Способа_оплаты join Услуги on Услуги.ID_Услуги = Договора.ID_Услуги";
            str.Open();
            SqlDataAdapter a = new SqlDataAdapter(query, str);
            DataSet ds = new DataSet();
            a.Fill(ds);
            dataGridView2.DataSource = ds.Tables[0];
            str.Close();
            panel1.Visible = false;
            panel2.Visible = true;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            String query = "SELECT * from Клиенты";
            str.Open();
            SqlDataAdapter a = new SqlDataAdapter(query, str);
            DataSet ds = new DataSet();
            a.Fill(ds);
            dataGridView3.DataSource = ds.Tables[0];
            str.Close();
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = true;
            panel4.Visible = false;
            panel5.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
                String query = "SELECT ID_сотрудника,Должности.Должность,ФИО,Дата_Рождения,Телефон,Серия_номер_паспорта,Логин,Пароль from Сотрудники join Должности on Должности.ID_Должности = Сотрудники.ID_Должности";
                str.Open();
                SqlDataAdapter a = new SqlDataAdapter(query, str);
                DataSet ds = new DataSet();
                a.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                str.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("У сотрудника активный договор!");
            }
           
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
        }

        private void Direktor_Load(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = true;
            panel5.Visible = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel1.Visible = true;
        }

        private void Direktor_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();
            Application.Exit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str.Open();
            SqlCommand command = new SqlCommand($"DELETE Сотрудники WHERE ID_Сотрудника = "+id+"", str);        
            command.ExecuteNonQuery();
            MessageBox.Show("Сотрудник успешно удалён!");
            str.Close();
            button3_Click(sender, e);
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }


        string Dolznost = "";
        string FIO = "";
        string Date_birthday = "";
        string Number = "";
        string passport = "";
        string Login = "";
        string Parol = "";
        private void button7_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = true;


            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str.Open();
            SqlCommand command = new SqlCommand("select ID_сотрудника, ID_Должности,ФИО,Дата_Рождения,Телефон,Серия_номер_паспорта,Логин,Пароль from Сотрудники WHERE ID_Сотрудника = " + id+"", str);        
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                Dolznost = reader[1].ToString();
                FIO = reader[2].ToString();
                Date_birthday = reader[3].ToString();
                Number = reader[4].ToString();
                passport = reader[5].ToString();
                Login = reader[6].ToString();
                Parol = reader[7].ToString();

            }
            reader.Close();
            str.Close();
            comboBox2.Text = Dolznost;
            textBox6.Text = FIO;
            dateTimePicker2.Text = Date_birthday;
            textBox5.Text = Number;
            textBox9.Text = Login;
            textBox10.Text = Parol;
            textBox4.Text = passport;

        }

        private void button10_Click(object sender, EventArgs e)
        {
            panel5.Visible = false;
            panel1.Visible = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {


                string dolznost = comboBox1.Text;
                int id_dolznost;
                if (dolznost == "Менеджер")
                {
                    id_dolznost = 2;
                }
                else
                {
                    id_dolznost = 1;
                }
                SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
                str.Open();
                SqlCommand command = new SqlCommand($"insert into  Сотрудники(ID_Должности,ФИО,Дата_Рождения,Телефон,Серия_номер_паспорта,Логин,Пароль) values('" + id_dolznost + "', '" + textBox1.Text + "', '" + dateTimePicker1.Value + "', '" + textBox2.Text + "', '" + textBox3.Text + "', '" + textBox8.Text + "', '" + textBox7.Text + "');", str);
                command.ExecuteNonQuery();
                MessageBox.Show("Сотрудник успешно добавлен!");
                str.Close();

                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                button3_Click(sender, e);
            }
            else
            {
                MessageBox.Show("Заполните данные");
            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
            string dolznost = comboBox2.Text;
            int id_dolznost;
            if (dolznost == "Менеджер")
            {
                id_dolznost = 2;
            }
            else
            {
                id_dolznost = 1;
            }

            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str.Open();
            SqlCommand command = new SqlCommand($"update Сотрудники set ID_Должности='" + id_dolznost + "' , ФИО = '" + textBox6.Text + "',Дата_рождения='" + dateTimePicker2.Value + "',Телефон='" + textBox5.Text + "',Серия_номер_паспорта='" + textBox4.Text +"',Логин='" + textBox9.Text + "',Пароль='" + textBox10.Text + "' WHERE ID_Сотрудника = "+id+"", str);         
            command.ExecuteNonQuery();
            MessageBox.Show("Информация успешно изменена!");
            str.Close();
            button7_Click(sender, e);
            panel5.Visible = false;
        }
    }
}

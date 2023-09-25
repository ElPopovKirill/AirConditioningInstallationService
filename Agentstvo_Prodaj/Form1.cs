using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Agentstvo_Prodaj
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static int id_sotrydnika;
        public static int name_sotrydnika;


        private void button1_Click(object sender, EventArgs e)
        {

            string login = textBox1.Text;
            string parol = textBox2.Text;


            SqlConnection str = new SqlConnection(@"Data Source=WIN-LK1QRPRQTC6\SQLEXPRESS;Initial Catalog=Kyrsach;Integrated Security=True");
            str.Open();
            SqlCommand command = new SqlCommand("select Должности.Должность, ФИО , ID_Сотрудника from Сотрудники JOIN Должности on Должности.ID_Должности = Сотрудники.ID_Должности where(Логин='" + login + "' and Пароль='" + parol + "');", str);
            SqlDataReader reader = command.ExecuteReader();
            string fio = "";
            string dolznost = "";
           
            while (reader.Read())
            {
                fio = reader[1].ToString();
                dolznost = reader[0].ToString();
                id_sotrydnika = Convert.ToInt32(reader[2]);
            }
            reader.Close();
            str.Close();



            if (dolznost == "Директор")
            {
                Direktor direktor = new Direktor();
                direktor.Show();
                this.Hide();
            }
            else if (dolznost == "Менеджер")
            {
                Menejer menejer = new Menejer();
                menejer.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Нет такого пользователя ");
            }

        }

      

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();
            Application.Exit();
        }
    }
}

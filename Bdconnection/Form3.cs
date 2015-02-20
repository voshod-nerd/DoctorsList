using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.Sql;
namespace Bdconnection
{
    public partial class Form3 : Form
    {
       // nsi element; 

        public Form3()
        {
            InitializeComponent();
         

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
           

            if ((textBox1.Text == "") || (textBox5.Text == "") || (textBox6.Text == "")) { MessageBox.Show("Заполнены не все обязательные поля. \r Обязательные поля: Код врача,Фамилия и Имя Врача"); }
            else 
            {
                SqlConnection conect = new SqlConnection();
                conect.ConnectionString = Properties.Settings.Default.ConString;
                try { 
                conect.Open();
                SqlCommand comand = new SqlCommand();

              


                comand.Parameters.Add("@IDDOKT", SqlDbType.Int).Value = Convert.ToInt32(textBox1.Text);

                if (textBox2.Text != "")
                {
                    comand.Parameters.Add("@LPUKOD", SqlDbType.Int).Value = Convert.ToInt32(textBox2.Text);
                }
                else { comand.Parameters.Add("@LPUKOD", SqlDbType.Int).Value = DBNull.Value; }

                if (textBox3.Text != "")
                {
                    comand.Parameters.Add("@ID_PODR", SqlDbType.Int).Value = Convert.ToInt32(textBox3.Text);
                }
                else
                { comand.Parameters.Add("@ID_PODR", SqlDbType.Int).Value = DBNull.Value; }


                if (textBox4.Text != "")
                {
                    comand.Parameters.Add("@ID_OTD", SqlDbType.Int).Value = Convert.ToInt32(textBox4.Text);
                }
                else
                { comand.Parameters.Add("@ID_OTD", SqlDbType.Int).Value = DBNull.Value; }

                if (textBox5.Text != "")
                {   
                comand.Parameters.Add("@FAM", SqlDbType.NVarChar).Value = textBox5.Text;
                } 
                else
                { comand.Parameters.Add("@FAM", SqlDbType.NVarChar).Value = DBNull.Value; }
                    
                    
                if (textBox6.Text != "")
                {    
                comand.Parameters.Add("@IM", SqlDbType.NVarChar).Value =(textBox6.Text);
                } 
                else
                { comand.Parameters.Add("@IM", SqlDbType.NVarChar).Value = DBNull.Value; }

                if (textBox7.Text != "")
                {    
                comand.Parameters.Add("@OT", SqlDbType.NVarChar).Value = (textBox7.Text);
                }     
                else
                { comand.Parameters.Add("@OT", SqlDbType.NVarChar).Value = DBNull.Value; }
      
  
                comand.Parameters.Add("@DOKT", SqlDbType.NVarChar).Value = radioButton1.Checked;
                
                if (textBox8.Text != "") 
                {
                comand.Parameters.Add("@PRVS", SqlDbType.Int).Value = Convert.ToInt32(textBox8.Text);
                }  
                else
                { comand.Parameters.Add("@PRVS", SqlDbType.Int).Value = DBNull.Value; }
                
                //дата внесения врача в реестр
                comand.Parameters.Add("@DATE_VN", SqlDbType.Date).Value = DateTime.Today;

                comand.CommandText = "INSERT INTO spisok_vrach (IDDOKT,LPUKOD,ID_PODR,ID_OTD,FAM,IM,OT,DOKT,PRVS,DATE_VN) VALUES (@IDDOKT,@LPUKOD,@ID_PODR,@ID_OTD,@FAM,@IM,@OT,@DOKT,@PRVS,@DATE_VN)";
                comand.CommandType = CommandType.Text;
                comand.Connection = conect;
                   // comand.
                comand.Parameters[1].IsNullable = true;
                comand.Parameters[2].IsNullable = true;
                comand.Parameters[3].IsNullable = true;
                comand.Parameters[4].IsNullable = true;
                comand.Parameters[5].IsNullable = true;
                comand.Parameters[6].IsNullable = true;
                comand.Parameters[7].IsNullable = true;
                comand.Parameters[8].IsNullable = true;
              
              



                comand.ExecuteNonQuery();
                conect.Close();
                MessageBox.Show("Данные занесены успешно");
                    // очистка данных по врачу 
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                }
                catch (Exception dataAcseesError)
                { MessageBox.Show(dataAcseesError.Message); conect.Close(); }
 







            }



        }

        private void button2_Click(object sender, EventArgs e)
        {
          
            if ((textBox9.Text == "") | (textBox11.Text == "") | (comboBox1.Text=="") | (ComboBoxKodandNamedDoctor.Text == "") )
            {
                MessageBox.Show("Одно или несколько из обзятательных полей не заполнено. Все поля обязательны к заполнению");
            }
            else 
              {
                SqlConnection conect = new SqlConnection();
                conect.ConnectionString = Properties.Settings.Default.ConString;
                try
                {
                    conect.Open();
                    SqlCommand comand = new SqlCommand();
                    comand.Parameters.Add("@N_SERT", SqlDbType.NVarChar).Value = (textBox9.Text);
                    if (textBox10.Text != "")
                    {
                        comand.Parameters.Add("@REG_NUM", SqlDbType.NVarChar).Value = (textBox10.Text);
                    }
                    else 
                    {
                        comand.Parameters.Add("@REG_NUM", SqlDbType.NVarChar).Value = DBNull.Value;
                    }   
                        
                        comand.Parameters.Add("@DATE_END", SqlDbType.Date).Value = Convert.ToDateTime(textBox11.Text);
                   
                    comand.Parameters.Add("@PRVS_S", SqlDbType.NVarChar).Value = comboBox1.Text;
                    nsi elem = new nsi();
                    elem.LoadNsi();
                    int t = elem.GetNumberSpec(comboBox1.Text);
                    comand.Parameters.Add("@PRVS", SqlDbType.Int).Value = elem.GetNumberSpec(comboBox1.Text);
                    comand.Parameters.AddWithValue("@DATEADD", DateTime.Today);

                    ////////////////////////////////////////////////// ввод кода  доктора 
                    string cod = "";
                    for (int i = 0; i <= 5; i++) { cod = cod + ComboBoxKodandNamedDoctor.Text[i];}
                    comand.Parameters.Add("@IDDOKT", SqlDbType.Int).Value = Convert.ToInt32(cod);
                    /////////////////////////////////////////////
                    comand.CommandText = "INSERT INTO sertif (N_SERT,REG_NUM,DATE_END,PRVS,PRVS_S,IDDOKT,DATEADD) VALUES (@N_SERT,@REG_NUM,@DATE_END,@PRVS,@PRVS_S,@IDDOKT,@DATEADD)";
                    comand.Connection = conect;
                    comand.ExecuteNonQuery();
                    MessageBox.Show("Сертификат успешно добавлен");
                    
                    // очистака полей соотвествующих сертификату
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox11.Text = "";
                   
                    comboBox1.Text = "";

                   
                }
                catch (Exception DataAcsess) { MessageBox.Show(DataAcsess.Message); }
 

              }

        }

        private void Form3_Load(object sender, EventArgs e)
        {
            // Загрузка НСИ по специальностям врачей

         
             nsi elem = new nsi();
           string temp = elem.LoadNsi();

           if (temp == "")
           {
               List<string> ListAllSpec = new List<string>();
               elem.GetAllSpec(ListAllSpec);
               for (int i = 0; i < ListAllSpec.Count; i++)
               {
                   comboBox1.Items.Add(ListAllSpec[i]);
               }
           }
           else { MessageBox.Show(temp); }
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form1 form = new Form1();
            foreach (Form x in Application.OpenForms) 
            {
                if (x is Form1) {form = (Form1)x;  }
            }
            form.UpdateData();


        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ComboBoxKodandNamedDoctor_MouseClick(object sender, MouseEventArgs e)
        {

            try
            {
                SqlConnection conect = new SqlConnection();
                conect.ConnectionString = Properties.Settings.Default.ConString;
                
                DataTable source = new DataTable();
                conect.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM spisok_vrach", conect);
               
                adapter.Fill(source);
                conect.Close();

                foreach ( DataRow x in source.Rows) {

                    ComboBoxKodandNamedDoctor.Items.Add(x[0] + " - " + x[4] + " " + x[5] + x[6]);
                }

            }
            catch (Exception xxx )
            { MessageBox.Show(xxx.Message); }


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

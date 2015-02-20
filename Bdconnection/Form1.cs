using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.Sql;
using System.IO;

namespace Bdconnection
{
    public partial class Form1 : Form
    {
       
        string connection = Properties.Settings.Default.ConString;
        DataTable tableVr = new DataTable();
        DataTable sertif = new DataTable();
        SqlConnection sqlcon = new SqlConnection();
        DataTable temp = new DataTable();

        SqlDataAdapter adapterSertif = new SqlDataAdapter();
        SqlDataAdapter adaperVr = new SqlDataAdapter();

        

        public Form1()
        {
            InitializeComponent();
        }
        //начальная инициализация формы + подключение к базе данных и экспорт данных
        private void Form1_Load(object sender, EventArgs e)
        {
            // востановление сохраненных настроек окна
            this.Width = Properties.Settings.Default.FormWidth;
            this.Height = Properties.Settings.Default.FormHeight;
            // получение строки подключения
            sqlcon.ConnectionString = connection;
            // подключение к базе данных 
            try
            {
                sqlcon.Open();
              //  MessageBox.Show("Подключение успешно");
            }
            catch (SqlException e1) 
            { 
                MessageBox.Show("Чтото не так :-(( "+e1.Message);
                Application.Exit();
            }
           
            // получение данных 
            
            //SqlDataAdapter adapterSertif = new SqlDataAdapter("SELECT * From sertif", sqlcon);
            adaperVr.SelectCommand = new SqlCommand("SELECT * FROM spisok_vrach", sqlcon);
           adapterSertif.SelectCommand = new SqlCommand("SELECT * From sertif", sqlcon);
           // SqlDataAdapter adaperVr = new SqlDataAdapter("SELECT * FROM spisok_vrach",sqlcon);
         //  adaperVr.select


            adapterSertif.Fill(sertif);
            adaperVr.Fill(tableVr);
            dataGridView1.DataSource =tableVr;
            temp = sertif.Clone();

            DataRowCollection  col = sertif.Rows;
            foreach (DataRow row in col) 
            {
                temp.ImportRow(row);
            }

            dataGridView2.DataSource = temp;
            podsvetka.podsv(dataGridView1, sertif);
            dataGridView1.Columns[0].HeaderText = "Код врача";
            dataGridView1.Columns[1].HeaderText = "Код ЛПУ";
            dataGridView1.Columns[2].HeaderText = "Код подразделения";
            dataGridView1.Columns[3].HeaderText = "Код отделения";
            dataGridView1.Columns[4].HeaderText = "Фамилия врача";
            dataGridView1.Columns[5].HeaderText = "Имя врача";
            dataGridView1.Columns[6].HeaderText = "Отчество врача";
            dataGridView1.Columns[7].HeaderText = "Признак врач (да/нет)/средний мед.персонал(FALSE)";
            dataGridView1.Columns[8].HeaderText = "Специальность врача";
            dataGridView1.Columns[9].HeaderText = "Дата внесения в реестр";

            dataGridView2.Columns[0].HeaderText = "Номер сертификата";
            dataGridView2.Columns[1].HeaderText = "Регистрационный номер сертификата";
            dataGridView2.Columns[2].HeaderText = "Дата окончания действия сертификата";
            dataGridView2.Columns[3].HeaderText = "Код специальности";
            dataGridView2.Columns[4].HeaderText = "Название специальности";
            dataGridView2.Columns[5].HeaderText = "Код врача,которму принадлежит сертификат";


            sqlcon.Close();
            label3.Text = "Врачей в базе  " + tableVr.Rows.Count.ToString() + ", и  сертификатов " + sertif.Rows.Count.ToString();
          //  label4.Text = "Всего  сертификатов = " + sertif.Rows.Count.ToString(); 

        }
        
        private void button1_Click(object sender, EventArgs e)
        {
      
           

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        // перерисвока красных линий при нажатии на ячейку..
        // то есть при смене строки..
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int value = (int)dataGridView1.Rows[e.RowIndex].Cells[0].Value;

                temp.Clear();


                SqlDataAdapter dat = new SqlDataAdapter("SELECT * FROM sertif where IDDOKT=" + value.ToString(), sqlcon);
                sqlcon.Open();
                dat.Fill(temp);
                sqlcon.Close();


            }
            catch (Exception w) { }
            // подсветка сертифкатов 
            podsvetka.podcvetkaSertif(this.dataGridView2);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {   //запись настроек
            Properties.Settings.Default.FormWidth = this.Width;
            Properties.Settings.Default.FormHeight=this.Height;
            //сохранение настроек
            Properties.Settings.Default.Save();
        }

        private void dataGridView1_Sorted(object sender, EventArgs e)
        {
            podsvetka.podsv(dataGridView1, sertif);

        }
        // поиск по коду врача
        private void button1_Click_1(object sender, EventArgs e)
        {
           try {
                
                 string   selectString =
                          tableVr.Columns[0].Caption + " ='" + textBox2.Text.Trim() + "'";
                



                DataRowCollection allRows =
                    ((DataTable)dataGridView1.DataSource).Rows;

                DataRow[] searchedRows =
                    tableVr.Select(selectString);

                int rowIndex = allRows.IndexOf(searchedRows[0]);

                dataGridView1.CurrentCell =
                    dataGridView1[0, rowIndex];
            }
            catch (Exception exp) { MessageBox.Show("Неверные данные"); } 
        }

       

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked) { checkBox1.Checked = false;  }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) { checkBox2.Checked = false;  }
        }



        // действие происходящие при перерисовке формы.. 
        // то есть перерисовка datagridview2,1 и их формы
        private void Form1_Resize(object sender, EventArgs e)
        {
            int y = panel2.Width;
            int y1 = dataGridView1.Location.Y;
            dataGridView1.Location = new Point(25, y1);
            dataGridView1.Width = y - 25 - 25;

            // по части сертификатов \
            int yy = panel3.Location.Y;
            panel3.Height = this.Height - yy-80;
            // по части сертификатов ширины
            panel3.Width = this.Width;
            // 
            dataGridView2.Height = panel3.Height;
            //
            int ys = panel3.Width;
            int y4 = dataGridView2.Location.Y;
            dataGridView2.Location = new Point(25, y4);
            dataGridView2.Width = ys - 25 - 25;

        }

        // экспорт врачей.. 
        private void button2_Click(object sender, EventArgs e)
        {
            Form2 nnn = new Form2();
            nnn.ShowDialog();
        }

      

// фильтрация по тексту в text1box в зависимости от отмеченных сheckbox-ов. 
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
       
            if (checkBox1.Checked) 
            {
                tableVr.DefaultView.RowFilter = "[FAM] LIKE '" + textBox1.Text.ToLower()+"*'";
            }
            if (checkBox2.Checked)
            {
                tableVr.DefaultView.RowFilter = "[IM] LIKE '" + textBox1.Text.ToLower()+"*'";
               
            }
           


        }

     

        private void добавитьСведеньяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 form = new Form3();
            form.ShowDialog();
        }


        public  void UpdateData() 
        {
            try 
               {
                   sqlcon.Open();
                   sertif.Clear();
                   tableVr.Clear();
                   adaperVr.Fill(tableVr);
                   adapterSertif.Fill(sertif);
                   sqlcon.Close();
                   podsvetka.podsv(dataGridView1,sertif);
               }
            catch (Exception dataEroror) { MessageBox.Show(dataEroror.Message); }

        }

        private void выгрузкаВЭксельToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Configure save file dialog box
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "Список врачай и сертификатов"; // Default file name
            dlg.DefaultExt = ".xls"; // Default file extension
            dlg.Filter = "Text documents (.xls)|*.xls"; // Filter files by extension
            dlg.ShowDialog();
            if (dlg.FileName != "")
            {
                StreamWriter write = new StreamWriter(dlg.FileName);
                write.WriteLine();
                write.Close();

                Ex exportExcel = new Ex();
                exportExcel.WriteExcel(dlg.FileName, tableVr, sertif);
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UpdateVrach();
            UpdateSertif();
        }
    



        private void UpdateVrach() 
        {
            try
            {
                sqlcon.Open();
                // Create the UpdateCommand.


                SqlCommand command = new SqlCommand(
                      "UPDATE spisok_vrach SET IDDOKT=@IDDOKT, LPUKOD = @LPUKOD, ID_PODR=@ID_PODR,ID_OTD=@ID_OTD, " +
                      "FAM=@FAM,IM=@IM,OT=@OT,DOKT=@DOKT,PRVS=@PRVS,DATE_VN=@DATE_VN " +
                 " WHERE IDDOKT=@oldIDDOKT", sqlcon);


                SqlParameter parameter = command.Parameters.Add(
        "@oldIDDOKT", SqlDbType.Int, 32, "IDDOKT");
                parameter.SourceVersion = DataRowVersion.Original;



                // Add the parameters for the UpdateCommand.
                command.Parameters.Add("@IDDOKT", SqlDbType.Int, 32, "IDDOKT");
                command.Parameters.Add("@LPUKOD", SqlDbType.Int, 32, "LPUKOD");
                command.Parameters.Add("@ID_PODR", SqlDbType.Int, 32, "ID_PODR");
                command.Parameters.Add("@ID_OTD", SqlDbType.Int, 32, "ID_OTD");
                command.Parameters.Add("@FAM", SqlDbType.NVarChar, 25, "FAM");
                command.Parameters.Add("@IM", SqlDbType.NVarChar, 25, "IM");
                command.Parameters.Add("@OT", SqlDbType.NVarChar, 25, "OT");
                command.Parameters.Add("@DOKT", SqlDbType.Bit, 1, "DOKT");
                command.Parameters.Add("@PRVS", SqlDbType.Int, 32, "PRVS");
                command.Parameters.Add("@DATE_VN", SqlDbType.Date,1, "DATE_VN");
                // command.Parameters.sou = DataRowVersion.Original;


                //    .UpdateCommand = command;
                adaperVr.UpdateCommand = command;



                adaperVr.Update(tableVr.Select("", "", DataViewRowState.ModifiedCurrent));
                sqlcon.Close();

            }
            catch (Exception error) { MessageBox.Show(error.Message); sqlcon.Close(); }
        
        
        
        }
        private void UpdateSertif() 
        {
            sqlcon.Open();
            try 
            {
               SqlCommand command = new SqlCommand(
        "UPDATE sertif SET N_SERT = @N_SERT, REG_NUM = @REG_NUM,DATE_END=@DATE_END,PRVS=@PRVS,PRVS_S=@PRVS_S " +
        "WHERE N_SERT = @oldN_SERT  AND IDDOKT = @IDDOKT",sqlcon);

               SqlParameter parameter = command.Parameters.Add(
      "@oldN_SERT", SqlDbType.NVarChar, 25, "N_SERT");
            // SqlParameter parameter1 = command.Parameters.Add(
              //      "@OLDREGNUM", SqlDbType.NVarChar, 25, "REG_NUM");
           //  SqlParameter parameter2 = command.Parameters.Add(
             //   "@OLDIDDOKT", SqlDbType.Int, 25, "IDDOKT");
              
                parameter.SourceVersion = DataRowVersion.Original;

               command.Parameters.Add("@N_SERT", SqlDbType.NVarChar, 25, "N_SERT");
               command.Parameters.Add("@REG_NUM", SqlDbType.NVarChar, 25, "REG_NUM");
               command.Parameters.Add("@DATE_END", SqlDbType.Date, 1, "DATE_END");
               command.Parameters.Add("@PRVS_S", SqlDbType.NVarChar,255, "PRVS_S");
               command.Parameters.Add("@PRVS", SqlDbType.Int, 32, "PRVS");
               command.Parameters.Add("@IDDOKT", SqlDbType.Int, 32, "IDDOKT");
              
               adapterSertif.UpdateCommand = command;
               adapterSertif.Update(temp);
               sqlcon.Close();
               UpdateData();
            
            }
            catch (Exception error) { MessageBox.Show(error.Message); }


        }

       

       

        private void CodeDoctors_Click(object sender, EventArgs e)
        {
            // Configure save file dialog box
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "Список кодов врачей"; // Default file name
            dlg.DefaultExt = ".xls"; // Default file extension
            dlg.Filter = "Text documents (.xls)|*.xls"; // Filter files by extension
            dlg.ShowDialog();
            if (dlg.FileName != "")
            {
                StreamWriter write = new StreamWriter(dlg.FileName);
                write.WriteLine();
                write.Close();

                Ex exportExcel = new Ex();
                exportExcel.WriteExcel(dlg.FileName, tableVr, sertif);
            }


        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void экспортВXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExToXML n = new ExToXML();

            n.GetData();
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
           
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                int kr = dataGridView1.CurrentRow.Index;
                int value = (int)dataGridView1.Rows[kr].Cells[0].Value;

                temp.Clear();


                SqlDataAdapter dat = new SqlDataAdapter("SELECT * FROM sertif where IDDOKT=" + value.ToString(), sqlcon);
                sqlcon.Open();
                dat.Fill(temp);
                sqlcon.Close();


            }
            catch (Exception w) { }
            // подсветка сертифкатов 
            podsvetka.podcvetkaSertif(this.dataGridView2);
        }

      
        
      

    }
}

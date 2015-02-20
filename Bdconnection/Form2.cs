using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;


namespace Bdconnection
{
    public partial class Form2 : Form
    {
        DataSet excel_vrachi = new DataSet();
        DataSet nsiSpecDoctros = new DataSet();
        public Form2()
        {
            InitializeComponent();
        }
        public string[] RetFamImOt(string text) 
        {
      
            string[] split = text.Split(new Char[] { ' '});

            return split;
        
        
        
        }

        public void exportSertif() 
        {
            SqlConnection sqlcon = new SqlConnection();
            sqlcon.ConnectionString = Properties.Settings.Default.ConString;


            try
            {

                for (int i = 2; i < excel_vrachi.Tables[0].Rows.Count; i++)
                {    
                    try {
                    var uvolen = excel_vrachi.Tables[0].Rows[i][0];
                    if ((Convert.ToString(uvolen) == ""))
                    {
                        sqlcon.Open();
                        SqlCommand sqlcom = new SqlCommand();

                        string SQLquery = "INSERT INTO sertif (IDDOKT,N_SERT,REG_NUM,DATE_END,PRVS,PRVS_S) VALUES (@IDDOKT,@N_SERT,@REG_NUM,@DATE_END,@PRVS,@PRVS_S)";
                        sqlcom.Parameters.Add("@IDDOKT", SqlDbType.Int).Value = excel_vrachi.Tables[0].Rows[i][1];
                        sqlcom.Parameters.Add("@N_SERT", SqlDbType.NVarChar).Value = excel_vrachi.Tables[0].Rows[i][5] + " " + excel_vrachi.Tables[0].Rows[i][6];
                        if (excel_vrachi.Tables[0].Rows[i][8] != null)
                        {
                            sqlcom.Parameters.Add("@REG_NUM", SqlDbType.NVarChar).Value = excel_vrachi.Tables[0].Rows[i][7];
                        }
                        else { sqlcom.Parameters.Add("@REG_NUM", SqlDbType.NVarChar).Value = ""; }

                        var dt = DateTime.FromOADate((double)excel_vrachi.Tables[0].Rows[i][9]);
                        
                        sqlcom.Parameters.Add("@DATE_END", SqlDbType.Date).Value = dt;
                        sqlcom.Parameters.Add("@PRVS_S", SqlDbType.NVarChar).Value = excel_vrachi.Tables[0].Rows[i][4];

                        string spec = "MSPNAME = '" + excel_vrachi.Tables[0].Rows[i][4] + "'";
                        DataRow[] findrows = nsiSpecDoctros.Tables[0].Select(spec);
                        if (findrows.Length != 0)
                        {
                            DataRow find = findrows[0];
                            object[] msvalue = find.ItemArray;
                            int numberspec = Convert.ToInt32(msvalue[2]);
                            sqlcom.Parameters.Add("@PRVS", SqlDbType.Int).Value = numberspec;
                        }
                        else { sqlcom.Parameters.Add("@PRVS", SqlDbType.Int).Value = -1; }
                        sqlcom.CommandType = CommandType.Text;
                        sqlcom.CommandText = SQLquery;
                        sqlcom.Connection = sqlcon;
                      
                        sqlcom.ExecuteNonQuery();
                        sqlcon.Close();
                    }
                    }
                    catch (Exception error1) { sqlcon.Close(); }
                }


            }
            catch (Exception error) {  }
        
        } 


        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog n = new OpenFileDialog();
            n.ShowDialog();
            if (n.FileName == "") { } else 
            {
                Ex ecclass = new Ex();
             excel_vrachi =  ecclass.FileExcelToDataSet(n.FileName);
             dataGridView1.DataSource= excel_vrachi.Tables[0];
            
            }

         
           
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int kolzap = 0;
            SqlConnection sqlcon = new SqlConnection(Properties.Settings.Default.ConString);
            /// Экспорт Врачей
            try 
            {
                
                
                for (int i = 2; i < excel_vrachi.Tables[0].Rows.Count; i++)
                {

                    string sqlComString = "";
                    try

                    {

                        SqlCommand command = new SqlCommand();
                        sqlComString = ""; 
                        var uvolen=excel_vrachi.Tables[0].Rows[i][0]; 
                        if ((Convert.ToString(uvolen)== ""))
                        {
                            sqlcon.Open();
                           
                            string[] m = RetFamImOt((string)excel_vrachi.Tables[0].Rows[i][2]);
                         if (m.Length == 2)
                            {
                                sqlComString = "INSERT INTO spisok_vrach (IDDOKT,FAM,IM) VALUES ( @IDDOKT, @FAM, @IM) ";
                                command.Parameters.Add("@FAM", SqlDbType.NVarChar).Value = m[0];
                                command.Parameters.Add("@IM", SqlDbType.NVarChar).Value = m[1];
                            }
                            if (m.Length == 3) { 
                              sqlComString = "INSERT INTO spisok_vrach (IDDOKT,FAM,IM,OT) VALUES ( @IDDOKT, @FAM, @IM, @OT) ";
                              command.Parameters.Add("@FAM",SqlDbType.NVarChar).Value=m[0];
                                command.Parameters.Add("@IM",SqlDbType.NVarChar).Value=m[1];
                              command.Parameters.Add("@OT",SqlDbType.NVarChar).Value=m[2];
                                
                            }
                                command.Parameters.Add("@IDDOKT",SqlDbType.Int).Value=excel_vrachi.Tables[0].Rows[i][1];
                              //  command.Parameters.Add("PRVS", SqlDbType.Int).Value = ;
                              //  string spec = "MSPNAME = '"+excel_vrachi.Tables[0].Rows[i][4]+"'";
                               // DataRow[] findrows = nsiSpecDoctros.Tables[0].Select(spec);
                               // DataRow find = findrows[0];
                              //  object[] msvalue = find.ItemArray;
                          // int  numberspec =  Convert.ToInt32(msvalue[2]);
                            //   command.Parameters.Add("PRVS", SqlDbType.Int).Value = numberspec;
                            
                            
                            
                            command.CommandType=CommandType.Text;
                                command.CommandText = sqlComString;
                                command.Connection = sqlcon;
                            command.ExecuteNonQuery();
                            kolzap++;
                            sqlcon.Close();


                        }
                    }
                    catch (Exception exs) { sqlcon.Close();}
                } 

            
            }
            catch (Exception exp) {}
            sqlcon.Close();
            MessageBox.Show("Врачи занесены");
            MessageBox.Show(kolzap.ToString());

            exportSertif();
            MessageBox.Show("Сертификаты тоже");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog opennsi = new OpenFileDialog();
            opennsi.ShowDialog();
            if (opennsi.FileName != "") 
            {
                StreamReader reader = new StreamReader(opennsi.FileName);
                nsiSpecDoctros.ReadXml(reader);
              //  dataGridView1.DataSource = nsiSpecDoctros.Tables[0];
            
            
            
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Xml;


namespace Bdconnection
{
    class ExToXML
    {
        SqlConnection conns = new SqlConnection();

        public void Ex() 
        {
            SqlConnection coon = new SqlConnection();
            coon.ConnectionString = Properties.Settings.Default.ConString;
            SqlDataAdapter adap = new SqlDataAdapter();
            SqlCommand com = new SqlCommand();
            com.CommandText = "SELECT * FROM spisok_vrach";
            com.Connection = coon;
            adap.SelectCommand = com;

            // отурытие 
            coon.Open();
            DataTable spVr = new DataTable();
            adap.Fill(spVr);

              
            foreach (DataRow x in spVr.Rows) 
            { 
            
            }










        
        
        }
        public void GetData()
        {

            conns.ConnectionString = Properties.Settings.Default.ConString;
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM spisok_vrach", conns);
            conns.Open();
            DataTable exprt = new DataTable();

            adapter.Fill(exprt);

            XmlDocument doc = new XmlDocument();
            doc.CreateXmlDeclaration("1.0", "windows-1251", "yes");

            XmlNode packet = doc.CreateElement("packet");


            int zap = 0;
            foreach (DataRow x in exprt.Rows)
            {
                zap++;

                XmlNode ZAP = doc.CreateElement("ZAP");
                XmlNode N_ZAP = doc.CreateElement("N_ZAP");
                N_ZAP.InnerText = zap.ToString();
                XmlNode IDDOCT = doc.CreateElement("IDDOCT");
                XmlNode LPUKOD = doc.CreateElement("LPUKOD");
                IDDOCT.InnerText = ((int)x[0]).ToString();
                if (x[1] == DBNull.Value) { LPUKOD.InnerText = ""; }
                else
                {
                    LPUKOD.InnerText = Convert.ToString(x[1]);
                }

                XmlNode FAM = doc.CreateElement("FAM");
                XmlNode IM = doc.CreateElement("IM");
                XmlNode OT = doc.CreateElement("OT");
                XmlNode DOKT = doc.CreateElement("DOKT");
                if (x[4] == DBNull.Value) { FAM.InnerText = ""; }
                else
                {
                    FAM.InnerText = ((string)x[4]);
                }

                if (x[5] == DBNull.Value) { IM.InnerText = ""; }
                else
                {
                    IM.InnerText = (string)x[5];
                }
                if (x[6] == DBNull.Value) { OT.InnerText = ""; }
                else
                {
                    OT.InnerText = ((string)x[6]);
                }
                if (x[7] == DBNull.Value) { DOKT.InnerText = "0"; }
                else
                {
                    string temp= ((bool)x[7]).ToString();
                    if (temp=="False") {DOKT.InnerText="0";}
                    if (temp == "True")
                    {
                        DOKT.InnerText = "1";
                        }
                }
                ZAP.AppendChild(N_ZAP);
                ZAP.AppendChild(IDDOCT);
                ZAP.AppendChild(LPUKOD);
                ZAP.AppendChild(FAM);
                ZAP.AppendChild(IM);
                ZAP.AppendChild(IM);
                ZAP.AppendChild(OT);
                ZAP.AppendChild(DOKT);




                // подлючение к таблице сертификатов
                SqlDataAdapter sertifAdapter = new SqlDataAdapter("SELECT * FROM sertif WHERE IDDOKT=" + Convert.ToString(x[0]), conns);
                DataTable srt = new DataTable();
                sertifAdapter.Fill(srt);

               
                foreach (DataRow row in srt.Rows)
                {
                    DateTime d = new DateTime(2015,02,20);
                    if (((DateTime)row[2])>d)
                    {
                    XmlNode SERTIF = doc.CreateElement("SERTIF");
                  
                  

                    XmlNode N_SERT = doc.CreateElement("N_SERT");
                    XmlNode REG_NUM = doc.CreateElement("REG_NUM");
                    XmlNode DATE_END = doc.CreateElement("DATE_END");
                    XmlNode PRVS = doc.CreateElement("PRVS");


                    if (row[0] == DBNull.Value) { N_SERT.InnerText = ""; }
                    else
                    {
                        N_SERT.InnerText = (string)row[0];
                    }

                    if (row[1] == DBNull.Value) { REG_NUM.InnerText =""; }
                    else
                    {
                        REG_NUM.InnerText = (string)row[1];
                    }
                    DATE_END.InnerText = ((DateTime)row[2]).ToString("dd.MM.yyyy");

                    string prvs = ((int)row[3]).ToString();
                    //string curd=  Directory.GetCurrentDirectory();
                    //INI ini = new INI(curd+"\\set.ini");
                    //string temp=    ini.IniReadValue("V004-v015", prvs);

                    PRVS.InnerText = prvs;

                    SERTIF.AppendChild(N_SERT);
                    SERTIF.AppendChild(REG_NUM);
                    SERTIF.AppendChild(DATE_END);
                    SERTIF.AppendChild(PRVS);

                    ZAP.AppendChild(SERTIF);
                }
                }

                packet.AppendChild(ZAP);  
            }

           
            doc.AppendChild(packet);
            doc.Save("simple.xml");
            conns.Close();

        }

    }
}

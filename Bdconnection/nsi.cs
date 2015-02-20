using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Data;

namespace Bdconnection
{
    class nsi
    {
        List<string> listSpecStr = new List<string>();
        List<int> listSpecInt = new List<int>();

        private string Encode(string source, System.Text.Encoding from, System.Text.Encoding to)
{

byte[] encodedsource = from.GetBytes(source);
return to.GetString(encodedsource);
}
       public DataSet nsiSpecDoc = new DataSet();
       public  string LoadNsi() 
        {
            try
            {
                StreamReader reader = new StreamReader("V015.xml");
                nsiSpecDoc.ReadXml(reader);
            }
            catch (Exception a) { return a.Message; }
            return "";
        }

       public void GetAllSpec(List<string> listAllSpec) 
        {
          
            DataTable table = new DataTable();
            table = nsiSpecDoc.Tables[0];
            int kol = table.Rows.Count;
        for (int i=0;i<nsiSpecDoc.Tables[0].Rows.Count;i++) 
            {
                string tp = (string)nsiSpecDoc.Tables[0].Rows[i][5];
                if (tp == "") { } else {

            listAllSpec.Add((string)nsiSpecDoc.Tables[0].Rows[i][2]);
            listSpecStr.Add((string)nsiSpecDoc.Tables[0].Rows[i][2]);
            listSpecInt.Add(Convert.ToInt32((string)nsiSpecDoc.Tables[0].Rows[i][1]));
                }
        }


      
        
        }

        public int GetNumberSpec(string spec)   
        {
        
            for (int i = 0; i < nsiSpecDoc.Tables[0].Rows.Count; i++) 
            {
                listSpecStr.Add((string)nsiSpecDoc.Tables[0].Rows[i][2]);
                listSpecInt.Add(Convert.ToInt32((string)nsiSpecDoc.Tables[0].Rows[i][1]));
            
            }
            int pos = listSpecStr.IndexOf(spec);
            return listSpecInt[pos];
        
        
        }
    }
}

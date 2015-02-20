using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Drawing;

namespace Bdconnection
{
   static class podsvetka
    {
       // Подстветка врачей у которых или нет сертификатов или они просрочены
    static public void   podsv(DataGridView vrachi, DataTable sertif) 
       {
           for (int i = 0; i < vrachi.Rows.Count; i++) 
           {
               // получение списка сертификатов одного врача 
               int zn = (int)vrachi.Rows[i].Cells[0].Value;
           DataRow[] vseSertifVracha = sertif.Select("IDDOKT = '"+zn.ToString()+"'");
               int koldeistv =0;
               // узнаем текущую дату
                DateTime current =  DateTime.Today;
               
               foreach (DataRow x in vseSertifVracha) 
               {
                  
                   var temp = x[2];
                   if (temp != null)
                   {
                       DateTime dt = Convert.ToDateTime(x[2]);
                       if (((dt.CompareTo(current)) == 1)||((dt.CompareTo(current)) == 0)) { koldeistv++; }
                   }
               }
               // если нет действующих сертификатов то подсвечиваем строку
               if (koldeistv == 0) 
               {  vrachi.Rows[i].DefaultCellStyle.BackColor=Color.Red;}


           
           
           }
       }

       // подсветка просроченных сертификатов 
    static public void podcvetkaSertif(DataGridView sertifs) 
    {
        DateTime current = DateTime.Today;
        for (int i = 0; i < sertifs.Rows.Count; i++) 
        {
            var temp = sertifs.Rows[i].Cells[2].Value;
            if (temp != null)
            {
                DateTime dt = Convert.ToDateTime(sertifs.Rows[i].Cells[2].Value);
                if (((dt.CompareTo(current)) == -1)) { sertifs.Rows[i].DefaultCellStyle.BackColor = Color.Red; }
            }  
        
        
        
        
        }
    
    
    
    }
    
    }
}

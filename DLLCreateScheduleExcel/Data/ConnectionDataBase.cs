using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace DLLCreateScheduleExcel.Data
{
     class ConnectionDataBase
    {
        private SqlConnection conn;
        private string strCnx;
        private SqlDataReader dr;

        public void Connect()
        {
            this.strCnx = "Server=osas;Initial Catalog=DB_Tgestiona;user id=sa;pwd=x";
            this.conn = new SqlConnection(strCnx);
            this.conn.Open();
        }
         
        public void Query(string query)
        {
            SqlCommand comm = new SqlCommand(query, this.conn);
            this.dr = comm.ExecuteReader();

            for(int i = 0; i<dr.FieldCount; i++)
            {
                Console.WriteLine("{0} ", dr.GetName(i));
            }
            Console.WriteLine();

            while (dr.Read())
            {
                for(int i = 0; i<dr.FieldCount; i++)
                {
                    Console.WriteLine("{0} ", dr[i]);
                }
            }
            this.dr.Close();
            this.conn.Close();
        }

        //public void CloseConnection()
        //{
        //    this.dr.Close();
        //    this.conn.Close();
        //}
    }
}

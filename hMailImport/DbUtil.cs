using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace hMailImport
{
    class DbUtil
    {

        public static MySqlConnection
                 GetDBConnection()
        {
            string host = "192.168.100.3";
            int port = 3306;
            string database = "oscmail";
            string username = "osc";
            string password = "gedeon4j";

            #if DEBUG
            host = "192.168.56.102";
            port = 3306;
            database = "oscmail";
            username = "root";
            password = "root";
            #endif


            // Connection String.
            string connString = "Server=" + host + ";Database=" + database
                + ";port=" + port + ";User Id=" + username + ";password=" + password+ ";Charset=utf8;";

            MySqlConnection conn = new MySqlConnection(connString);

            return conn;
        }        

    }
}

using System;
using System.Diagnostics;
using System.Data.Odbc;
using System.Data;

namespace CostOfCapital
{
    public class Connection
    {

        //Connect to the database, create connectivity and command elements
        OdbcConnection dbConnect = new OdbcConnection();
        string myConnString = Properties.Settings.Default.#Connection;

        //Run query, return data reader
        public OdbcDataReader RunQuery(String SSQL)
        {

            OdbcCommand dbCommand = new OdbcCommand
            {
                Connection = dbConnect
            };


            try
            {
                OpenConn();
                dbCommand.CommandText = SSQL;
                OdbcDataReader myReader = dbCommand.ExecuteReader();

                return myReader;

            }
            catch (Exception e)
            {
                Debug.WriteLine("DB Reader Error:  " + e.Source + ", " + e.Message);
            }

            CloseConn();
            return null;
        }

        //Pass information to DB
        public void SendQuery(String SSQL)
        {

            OdbcCommand dbCommand = new OdbcCommand
            {
                Connection = dbConnect
            };

            try
            {
                OpenConn();
                dbCommand.CommandText = SSQL;
                dbCommand.ExecuteNonQuery();
                CloseConn();

            }
            catch (Exception e)
            {
                Debug.WriteLine("Pass to DB Error:  " + e.Source + ", " + e.Message);
            }

            CloseConn();

        }

        //Open Connection
        public void OpenConn()
        {

            if (dbConnect.State == ConnectionState.Closed)
            {
                try
                {
                    dbConnect.ConnectionString = myConnString;
                    dbConnect.Open();
                }
                catch (Exception e)
                {
                    Debug.WriteLine("Open Error:  " + e.Source + ", " + e.Message);
                }
            }
        }

        //Close connection
        public void CloseConn()
        {
            dbConnect.Close();
        }




    }
}

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_Watermark
{
    class Program
    {
        public static SqlConnection con;

        public class ProcessFile
        {
            public string ID { get; set; }
            public string sourcePath { get; set; }
            public string destinationPath { get; set; }
        }

        static void Main(string[] args)
        {
            con = new SqlConnection("Application Name=Excel.Watermark;Data Source=10.0.0.17;Network Library=dbmssocn;Initial Catalog=INTRA_DB;User ID=sa;Password=ift.2017;Workstation ID=INTRA.Service;");

            for (int i = 0; i < 19; i++)
            {
                i += Process();
                if (i != 18)
                    Thread.Sleep(3000);
            }
        }

        public static int Process()
        {
            ErrorHandler ErrorHandler = new ErrorHandler();

            int result = 0;

            try
            {
                FileHandler fileHandler = new FileHandler();

                List<ProcessFile> processFiles = GetFilesToProcess();

                foreach (ProcessFile file in processFiles)
                {
                    if (file.destinationPath == "")
                        result += fileHandler.ProcessFiles(file.sourcePath);
                    else
                        result += fileHandler.ProcessFiles(file.sourcePath, file.destinationPath);

                    MarkFileAsProcessed(file.ID);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.SendError("Run", ex.ToString());
                Console.WriteLine(ex.ToString());
            }

            return result;
        }

        /// <summary>
        /// Vrati subory, ktore treba spracovat
        /// </summary>
        /// <returns></returns>
        public static List<ProcessFile> GetFilesToProcess()
        {
            List<ProcessFile> results = new List<ProcessFile>();

            string query = String.Format("SELECT * FROM lst_WatermarkExcel WHERE RS='A'");
            SqlCommand cmd = new SqlCommand(query, con);
            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ProcessFile file = new ProcessFile();
                file.ID = reader["ID"].ToString();
                file.sourcePath = reader["Path"].ToString();
                file.destinationPath = reader["OutputPath"].ToString();
                results.Add(file);
            }
            con.Close();

            return results;
        }

        /// <summary>
        /// Oznaci subor ako spracovany
        /// </summary>
        public static void MarkFileAsProcessed(string ID)
        {
            string query = String.Format("UPDATE lst_WatermarkExcel SET RS='D' WHERE ID='{0}'", ID);
            SqlCommand cmd = new SqlCommand(query, con);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
}

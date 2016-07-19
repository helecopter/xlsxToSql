using System;
using Excel;
using System.IO;
using System.IO.Compression;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace xlsxToSql
{
    class Program
    {
        static void Main(string[] args)
        {
            ParserXlsx k = new ParserXlsx();
            k.OpenExcel("C:\\Users\\ZHENGST\\Desktop\\Book1.xlsx");

        }
    }

    class ZipParser
    {
        private string default_extract_path = @"C:\user\document\ZipRelease";

        //https://msdn.microsoft.com/zh-cn/library/system.io.filestream(v=vs.110).aspx
        public void copyFile(string source_path,string destination_path)
        {
            using (FileStream read_stream = File.Open(source_path, FileMode.Open))
            {
                if(!File.Exists(destination_path))
                {
                    using (FileStream write_stream = File.Create(destination_path))
                    {
                        

                        byte[] buffer = new byte[1024];
                        while(read_stream.Read(buffer,0,buffer.Length) > 0)
                        {
                            write_stream.Write(buffer, 0, buffer.Length);
                        }
                        
                    }
                }
            }
        }

        //https://msdn.microsoft.com/zh-cn/library/ms404280(v=vs.110).aspx
        public void extract_file(string source_path, string destination_path)
        {
            using (ZipArchive archive = ZipFile.OpenRead(source_path))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".xlsx"))
                    {
                        entry.ExtractToFile(Path.Combine(destination_path, entry.FullName));
                    }
                }
            }
        }
    }


    class ParserXlsx
    {
        private string Server_Name = "(local)";
        private string DB_Name = "Excel_DB";


        



        public void OpenExcel(String strFileName)
        {
            string connectStr = "Server=" + Server_Name + ";Database=" + DB_Name + ";Integrated Security=True";
            FileStream stream = null;
            SqlConnection connect = null;
            IExcelDataReader excelReader = null;

            try
            {
                stream = File.Open(strFileName, FileMode.Open, FileAccess.Read);
                byte[] b = new byte[10000];
                stream.Read(b, 0, (int)stream.Length);
                
                if (strFileName.EndsWith(".xls"))
                {
                    //读取.xls文件
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (strFileName.EndsWith(".xlsx"))
                {
                    //读取.xlsx文件
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                excelReader.IsFirstRowAsColumnNames = true;
                DataSet result = excelReader.AsDataSet();


                if (result.Tables.Count < 1)
                {
                    Console.WriteLine("excel has no sheet");
                }

                System.Data.DataTable sheet = result.Tables[0];
                if (sheet.Rows.Count <= 0)
                {
                    Console.WriteLine("the sheet has no data");
                }

                System.Data.DataTable tb = result.Tables[0];


                connect = new SqlConnection(connectStr);
                connect.Open();

                SqlCommand createCommand = new SqlCommand();
                SqlCommand insertCommand = new SqlCommand();
                SqlCommand ExistCommand = new SqlCommand();



                createCommand.CommandText = "CREATE TABLE " + tb.TableName + "(";

                foreach (DataColumn col in tb.Columns)
                {
                    createCommand.CommandText = createCommand.CommandText + col.ColumnName + " VARCHAR(100) " + ",";
                }
                createCommand.CommandText = createCommand.CommandText.Substring(0, createCommand.CommandText.Length - 1);
                createCommand.CommandText = createCommand.CommandText + ")";
                createCommand.Connection = connect;

                //判断所要建表是否已经存在
                ExistCommand.CommandText = "select * from sysobjects where id = object_id('" + DB_Name + ".Owner." + tb.TableName + "')";
                ExistCommand.Connection = connect;
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(ExistCommand);
                if (sqlDataAdapter != null)
                {
                    createCommand.ExecuteNonQuery();
                }
                
                int rowLineNum = tb.Rows.Count;
                insertCommand.Connection = connect;
                for (int i = 1; i < rowLineNum; i++)
                {
                    DataRow dr = tb.Rows[i];
                    insertCommand.CommandText = "INSERT INTO " + tb.TableName + "  VALUES(";
                    foreach (object col in dr.ItemArray)
                    {

                        insertCommand.CommandText = insertCommand.CommandText + "'" + col.ToString() + "' ,";
                    }
                    insertCommand.CommandText = insertCommand.CommandText.Substring(0, insertCommand.CommandText.Length - 1);
                    insertCommand.CommandText = insertCommand.CommandText + ")";
                    insertCommand.ExecuteNonQuery();
                }
                connect.Close();
                excelReader.Close();
                stream.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("出错信息::" + e.Message);
            }
        }
    }
}

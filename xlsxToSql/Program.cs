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
            ZipParser.test();
        }
    }

    class ZipParser
    {
        private static string default_extract_path = @"C:\Users\ZHENGST\Documents";
        private static string default_test_file = @"D:\PIERS_EXPORT_WEEKLY_10252014.zip";

        public static void test()
        {
            extract_file(default_test_file, default_extract_path);
        }

        //https://msdn.microsoft.com/zh-cn/library/system.io.filestream(v=vs.110).aspx
        public void copyFile(string source_path, string destination_path)
        {
            using (FileStream read_stream = File.Open(source_path, FileMode.Open))
            {
                if (!File.Exists(destination_path))
                {
                    using (FileStream write_stream = File.Create(destination_path))
                    {
                        byte[] buffer = new byte[1024];
                        while (read_stream.Read(buffer, 0, buffer.Length) > 0)
                        {
                            write_stream.Write(buffer, 0, buffer.Length);
                        }
                    }
                }
            }
        }

        //https://msdn.microsoft.com/zh-cn/library/ms404280(v=vs.110).aspx
        //use absolute path as path
        public static void extract_file(string source_path, string destination_path)
        {
            using (ZipArchive archive = ZipFile.OpenRead(source_path))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".csv") || entry.FullName.EndsWith(".xlsx") || entry.FullName.EndsWith(".xls")) 
                    {
                        string extract_file_path = Path.Combine(destination_path, entry.FullName);
                        if (!File.Exists(extract_file_path))
                        {
                            entry.ExtractToFile(extract_file_path);
                        }
                        ParserXlsx.OpenExcel(extract_file_path);
                    }
                }
            }
        }
    }


    class Encoding_Get
    {
        public static System.Text.Encoding GetType(FileStream fs)
        {
            byte[] Unicode = new byte[] { 0xFF, 0xFE, 0x41 };
            byte[] UnicodeBIG = new byte[] { 0xFE, 0xFF, 0x00 };
            byte[] UTF8 = new byte[] { 0xEF, 0xBB, 0xBF };
            Encoding reVal = Encoding.Default;

            BinaryReader r = new BinaryReader(fs, System.Text.Encoding.Default);
            int i;
            int.TryParse(fs.Length.ToString(), out i);
            byte[] ss = r.ReadBytes(i);
            if (IsUTF8Bytes(ss) || (ss[0] == UTF8[0] && ss[1] == UTF8[1] && ss[2] == UTF8[2]))
            {
                reVal = Encoding.UTF8;
            }
            else if (ss[0] == UnicodeBIG[0] && ss[1] == UnicodeBIG[1] && ss[2] == UnicodeBIG[2])
            {
                reVal = Encoding.BigEndianUnicode;
            }
            else if (ss[0] == Unicode[0] && ss[1] == Unicode[1] && ss[2] == Unicode[2])
            {
                reVal = Encoding.Unicode;
            }
            r.Close();
            return reVal;
        }
        
        private static bool IsUTF8Bytes(byte[] data)
        {
            int charByteCount = 1;
            byte curByte;
            foreach(byte element in data)
            {
                curByte = element;
                if (charByteCount == 1)
                {
                    if (curByte >= 0x80)
                    {
                        while (((curByte <<= 1) & 0x80) != 0)
                        {
                            charByteCount++;
                        }
                        if (charByteCount == 1 || charByteCount > 6)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    if ((curByte & 0xC0) != 0x80)
                    {
                        return false;
                    }
                    charByteCount--;
                }
                if (charByteCount > 1)
                {
                    throw new Exception("not expect byte formate");
                }
                return true;

            }
            return true;
        }
    }

    class csvReader
    {
        
        public static DataSet CreateCsvReader(FileStream stream,string file_name)
        {
            //Encoding encoding = Encoding_Get.GetType(stream);
            //StreamReader stream_reader = new StreamReader(stream, encoding);

            DataSet ds = new DataSet();
            StreamReader stream_reader = new StreamReader(stream);
            string strLine = null;
            bool flag = false;
            DataTable dt = new DataTable();
            int column_count = 0;
            string[] ary_line = null;
            while ((strLine = stream_reader.ReadLine()) != null)
            {
                ary_line = strLine.Split(',');
                if (!flag)
                {
                    flag = true;
                    column_count = ary_line.Length;
                    for (int i = 0; i < column_count; i++) 
                    {
                        DataColumn dc = new DataColumn(ary_line[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    DataRow row = dt.NewRow();
                    for (int j = 0; j < column_count; j++)
                    {
                        row[j] = ary_line[j];
                    }
                    dt.Rows.Add(row);
                    
                }
            }

            //string table_name = file_name.Split
            int endpos = file_name.LastIndexOf('.');
            int startpos = file_name.LastIndexOf('\\');
            string table_name = file_name.Substring(startpos + 1, endpos - startpos - 1);
            dt.TableName = table_name;
            ds.Tables.Add(dt);
            return ds;
        }
    }


    class ParserXlsx
    {
        private static string Server_Name = "(local)";
        private static string DB_Name = "Excel_DB";
        
        public static void OpenExcel(String strFileName)
        {
            string connectStr = "Server=" + Server_Name + ";Database=" + DB_Name + ";Integrated Security=True";
            FileStream stream = null;
            SqlConnection connect = null;
            IExcelDataReader excelReader = null;
            DataSet result = null;
            try
            {
                stream = File.Open(strFileName, FileMode.Open, FileAccess.Read);
                byte[] b = new byte[10000000];


                /*
                 * 处理xlsx文件或xls文件
                 */
                if (strFileName.EndsWith(".xlsx") || strFileName.EndsWith(".xls"))
                {
                    stream.Read(b, 0, (int)stream.Length);
                    if (strFileName.EndsWith(".xls"))
                    {
                        excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (strFileName.EndsWith(".xlsx"))
                    {
                        excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }


                    excelReader.IsFirstRowAsColumnNames = true;
                    result = excelReader.AsDataSet();

                    excelReader.Close();
                }
                /*
                 * 处理csv文件
                 */
                else if (strFileName.EndsWith(".csv"))
                {
                    result = csvReader.CreateCsvReader(stream,strFileName);
                }


                /*
                 * 判断数据表是否存在
                 */
                if (result.Tables.Count < 1)
                {
                    Console.WriteLine("excel has no sheet");
                }

                /*
                 * 判断数据集是否为空
                 */
                System.Data.DataTable sheet = result.Tables[0];
                if (sheet.Rows.Count <= 0)
                {
                    Console.WriteLine("the sheet has no data");
                }

                
                System.Data.DataTable tb = result.Tables[0];


                /*
                 * 建立数据库连接
                 */
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

                /*
                 * 判断所要建表是否已经存在
                 */
                ExistCommand.CommandText = "select * from sysobjects where id = object_id('" + DB_Name + ".Owner." + tb.TableName + "')";
                ExistCommand.Connection = connect;
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(ExistCommand);
                DataSet tempds = new DataSet();
                sqlDataAdapter.Fill(tempds, "ee");
                if ((tempds.Tables[0].Rows.Count == 0)) 
                {
                    throw new Exception("the table has already exists");
                }
                createCommand.ExecuteNonQuery();

                /*
                 * 插入数据
                 */
                int rowLineNum = tb.Rows.Count;
                insertCommand.Connection = connect;
                for (int i = 0; i < rowLineNum; i++)
                {
                    DataRow dr = tb.Rows[i];
                    insertCommand.CommandText = "INSERT INTO " + tb.TableName + "  VALUES(";
                    foreach (object col in dr.ItemArray)
                    {
                        string content = col.ToString().Replace("'", "''");
                        insertCommand.CommandText = insertCommand.CommandText + "'" + content + "' ,";
                    }
                    insertCommand.CommandText = insertCommand.CommandText.Substring(0, insertCommand.CommandText.Length - 1);
                    insertCommand.CommandText = insertCommand.CommandText + ")";
                    insertCommand.ExecuteNonQuery();
                }
                connect.Close();
                stream.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("error line"+ e.Source);
                Console.WriteLine("出错信息::" + e.Message);
            }
        }
    }
}

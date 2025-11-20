using System;
using System.IO;
using System.Data;
using System.Text;
using ExcelDataReader;

namespace FileReaderTrans
{
    public class FileReaderTrans
    {
        /// <summary>
        /// 讀取檔案轉換為DataSet
        /// </summary>
        /// <param name="filePath">完整檔案路徑</param>
        /// <param name="Message">回傳訊息</param>
        /// <returns></returns>
        /// 
        public static DataSet ReaderToDataSet(string filePath, string fileName, out string Message)
        {
            DataSet ds = new DataSet();

            DirectoryInfo difo = new DirectoryInfo(filePath);
            if (difo.Exists == false)
            {
                Message = $"檔案 [{filePath}] 不存在!";
                return ds;
            }
            if (File.Exists(@filePath + fileName))
            {
                //副檔名
                var extension = Path.GetExtension(@filePath + fileName).ToLower();

                using (var stream = new FileStream(@filePath + fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //判斷格式套用讀取方法
                    IExcelDataReader reader = null;
                    if (extension == ".xls")
                    {
                        //(" => XLS格式");
                        reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("big5")
                        });
                    }
                    else if (extension == ".xlsx")
                    {
                        //(" => XLSX格式");
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else if (extension == ".csv")
                    {
                        //(" => CSV格式");
                        reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("big5")
                        });
                    }
                    else if (extension == ".txt")
                    {
                        //(" => Text(Tab Separated)格式");
                        reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                        {
                            FallbackEncoding = Encoding.GetEncoding("big5"),
                            AutodetectSeparators = new char[] { '\t' }
                        });
                    }

                    //沒有對應產生任何格式
                    if (reader == null)
                    {
                        // MessageBox.Show($"未知的處理檔案：{extension}");
                        Message = $"未知的處理檔案類型：[{extension}]";
                        return ds;
                    }

                    using (reader)
                    {
                        ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            UseColumnDataType = false,
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                //設定讀取資料時是否忽略標題
                                UseHeaderRow = false
                            }
                        });
                    }
                }
                Message = "";
                return ds;
            }
            else
            {
                Message = $"檔案 [{filePath}] 不存在!";
                return ds;
            }
            // public static DataSet ReaderToDataSet(string filePath, out string Message)
            //if (File.Exists(filePath))
            //{
            //    //副檔名
            //    var extension = Path.GetExtension(filePath).ToLower();

            //    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            //    {
            //        //判斷格式套用讀取方法
            //        IExcelDataReader reader = null;
            //        if (extension == ".xls")
            //        {
            //            //(" => XLS格式");
            //            reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
            //            {
            //                FallbackEncoding = Encoding.GetEncoding("big5")
            //            });
            //        }
            //        else if (extension == ".xlsx")
            //        {
            //            //(" => XLSX格式");
            //            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //        }
            //        else if (extension == ".csv")
            //        {
            //            //(" => CSV格式");
            //            reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
            //            {
            //                FallbackEncoding = Encoding.GetEncoding("big5")
            //            });
            //        }
            //        else if (extension == ".txt")
            //        {
            //            //(" => Text(Tab Separated)格式");
            //            reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
            //            {
            //                FallbackEncoding = Encoding.GetEncoding("big5"),
            //                AutodetectSeparators = new char[] { '\t' }
            //            });
            //        }

            //        //沒有對應產生任何格式
            //        if (reader == null)
            //        {
            //            // MessageBox.Show($"未知的處理檔案：{extension}");
            //            Message = $"未知的處理檔案類型：[{extension}]";
            //            return ds;
            //        }

            //        using (reader)
            //        {
            //            ds = reader.AsDataSet(new ExcelDataSetConfiguration()
            //            {
            //                UseColumnDataType = false,
            //                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
            //                {
            //                    //設定讀取資料時是否忽略標題
            //                    UseHeaderRow = false
            //                }
            //            });
            //        }
            //    }
            //    Message = "";
            //    return ds;
            //}
            //else
            //{
            //    Message = $"檔案 [{filePath}] 不存在!";
            //    return ds;
            //}
        }
    }
}

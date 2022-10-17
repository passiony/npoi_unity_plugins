using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;

public class ExcelMenus : Editor
{
    static string SrcPath = "/Config/data.xlsx";

    [MenuItem("Tools/Read Excel")]
    public static void ReadExcel()
    {
        string path = Directory.GetCurrentDirectory() + SrcPath;
        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            var workbook = WorkbookFactory.Create(stream);
            ISheet sheet = workbook.GetSheetAt(0);

            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (IRow row in sheet)
                {
                    sb.Clear();
                    for (var i = 0; i < row.LastCellNum; i++)
                    {
                        var cell = row.GetCell(i);
                        if (cell != null)
                        {
                            sb.Append(cell);
                            sb.Append("\t");
                        }
                    }

                    Debug.Log(sb.ToString());
                }
            }
            catch (Exception e)
            {
                Debug.LogError(e);
                throw;
            }
        }
    }

    static bool GetKey(object[][] array, string key, string value)
    {
        for (int j = 0; j < array.Length; j++)
        {
            var txt = array[j][1].ToString();
            if (txt.Contains(value))
            {
                Debug.Log("value:" + value);
                array[j][0] = key;
                return true;
            }
        }

        return false;
    }

    public static object[][] ReadFromExcelFile(string filePath)
    {
        List<object[]> array = new List<object[]>();
        string extension = Path.GetExtension(filePath);
        IWorkbook wk;
        OPCPackage pkg = null;
        int num = 0;
        try
        {
            if (extension.Equals(".xls"))
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    wk = new HSSFWorkbook(fs);
                }
            }
            else
            {
                pkg = OPCPackage.Open(filePath, PackageAccess.READ_WRITE);
                wk = new XSSFWorkbook(pkg);
            }

            //读取当前表数据
            ISheet sheet = wk.GetSheetAt(0);
            int rowNum = sheet.LastRowNum;
            IRow row; //读取当前行数据

            for (int i = 0; i <= rowNum; i++)
            {
                num = i;
                row = sheet.GetRow(i); //读取当前行数据
                if (row != null)
                {
                    List<object> data = new List<object>();
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        //读取该行的第j列数据
                        var cell = row.GetCell(j);
                        string value = "";
                        if (cell != null)
                        {
                            value = cell.ToString();
                        }

                        data.Add(value);
                    }

                    array.Add(data.ToArray());
                }
            }
        }

        catch (Exception e)
        {
            //只在Debug模式下才输出
            Debug.LogError(e.Message + num);
            pkg?.Revert();
        }

        return array.ToArray();
    }

    public static void WriteToExcel(string filePath, object[][] data)
    {
        int num = 0;
        try
        {
            //创建工作薄  
            IWorkbook wb;
            string extension = System.IO.Path.GetExtension(filePath);
            //根据指定的文件格式创建对应的类
            if (extension.Equals(".xls"))
            {
                wb = new HSSFWorkbook();
            }
            else
            {
                wb = new XSSFWorkbook();
            }

            //创建一个表单
            ISheet sheet = wb.CreateSheet("Sheet0");
            int rowCount = data.Length;
            int columnCount = data[0].Length;

            for (int i = 0; i < rowCount; i++)
            {
                num = i;
                var row = sheet.CreateRow(i); //创建第i行
                for (int j = 0; j < columnCount; j++)
                {
                    var cell = row.CreateCell(j); //创建第j列
                    var line = data[i];
                    if (line.Length > j)
                    {
                        object obj = data[i][j];
                        SetCellValue(cell, obj);
                    }
                }
            }

            FileStream fs = File.OpenWrite(filePath);
            wb.Write(fs); //向打开的这个Excel文件中写入表单并保存。 
            fs.Flush();
            fs.Close();
        }
        catch (Exception e)
        {
            Debug.LogError(e.Message + num);
        }
    }

    public static void SetCellValue(ICell cell, object obj)
    {
        if (obj is int)
        {
            cell.SetCellValue((int)obj);
        }
        else if (obj is double)
        {
            cell.SetCellValue((double)obj);
        }
        else if (obj.GetType() == typeof(IRichTextString))
        {
            cell.SetCellValue((IRichTextString)obj);
        }
        else if (obj is string)
        {
            cell.SetCellValue(obj.ToString());
        }
        else if (obj is DateTime)
        {
            cell.SetCellValue((DateTime)obj);
        }
        else if (obj is bool)
        {
            cell.SetCellValue((bool)obj);
        }
        else
        {
            cell.SetCellValue(obj.ToString());
        }
    }
}
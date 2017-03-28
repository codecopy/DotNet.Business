/****************************************************************************************
 *功能实现：实现Excel导入导出功能 *******************************************************
 ****************************************************************************************/
namespace DotNet.Business.NPOIHelper
{
    #region 引用
    using System;
    using System.Data;
    using System.IO;
    using System.Text;
    using System.Web;
    using NPOI;
    using NPOI.SS.UserModel;
    using NPOI.HPSF;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.POIFS;
    using NPOI.Util;
    using NPOI.SS.Util;
    using System.Collections.Generic;
    #endregion
    #region 构造Excel导入导出类
    public class ExcelHelper
    {
        #region 单纯生成EXCEL
        /// <summary>
        /// 导出到Excel
        /// </summary>
        /// <param name="dtSource"></param>
        /// <param name="strFileName"></param>
        /// <remarks>NPOI认为Excel的第一个单元格是：(0，0)</remarks>
        public static void ExportEasy(DataTable dtSource, string strFileName)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            Sheet sheet = workbook.CreateSheet();

            //填充表头
            Row dataRow = sheet.CreateRow(0);
            foreach (DataColumn column in dtSource.Columns)
            {
                dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
            }
            //填充内容
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dtSource.Rows[i][j].ToString());
                }
            }
            //保存
            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(ms);
                    ms.Flush();
                    ms.Position = 0;
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
            sheet.Dispose();
            workbook.Dispose();
        }
        /// <summary>
        /// DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        public static MemoryStream Export(DataTable dtSource, string strHeaderText)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            Sheet sheet = workbook.CreateSheet();

            #region 右击文件 属性信息






            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();

                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }
            #endregion

            CellStyle dateStyle = workbook.CreateCellStyle();
            DataFormat format = workbook.CreateDataFormat();
            //dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            //取得列宽
            int[] arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;
            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式
                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet();
                    }

                    #region 表头及样式






                    {
                        Row headerRow = sheet.CreateRow(0);
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        CellStyle headStyle = workbook.CreateCellStyle();
                        //headStyle.Alignment = CellHorizontalAlignment.CENTER;
                        Font font = workbook.CreateFont();
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);

                        headerRow.GetCell(0).CellStyle = headStyle;

                        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1));
                        //headerRow.Dispose();
                    }
                    #endregion


                    #region 列头及样式






                    {
                        Row headerRow = sheet.CreateRow(1);


                        CellStyle headStyle = workbook.CreateCellStyle();
                        //headStyle.Alignment = CellHorizontalAlignment.CENTER;
                        Font font = workbook.CreateFont();
                        font.FontHeightInPoints = 10;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);


                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                            //设置列宽
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);

                        }
                        //headerRow.Dispose();
                    }
                    #endregion

                    rowIndex = 2;
                }
                #endregion
                #region 填充内容
                Row dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dtSource.Columns)
                {
                    Cell newCell = dataRow.CreateCell(column.Ordinal);

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String"://字符串类型






                            newCell.SetCellValue(drValue);
                            break;
                        case "System.DateTime"://日期类型
                            DateTime dateV;
                            DateTime.TryParse(drValue, out dateV);
                            newCell.SetCellValue(dateV);

                            newCell.CellStyle = dateStyle;//格式化显示






                            break;
                        case "System.Boolean"://布尔型






                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case "System.Decimal"://浮点型






                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case "System.DBNull"://空值处理






                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue("");
                            break;
                    }

                }
                #endregion
                rowIndex++;
            }
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;

                sheet.Dispose();
                workbook.Dispose();

                return ms;
            }
        }
        /// <summary>
        /// DataTable导出到Excel文件
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">保存位置</param>
        public static void Export(DataTable dtSource, string strHeaderText, string strFileName)
        {
            using (MemoryStream ms = Export(dtSource, strHeaderText))
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }
        /// <summary>
        /// 用于Web导出
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">文件名</param>
        public static void ExportByWeb(DataTable dtSource, string strHeaderText, string strFileName)
        {

            HttpContext curContext = HttpContext.Current;
            // 设置编码和附件格式
            curContext.Response.ContentType = "application/vnd.ms-excel";
            curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.Charset = "";
            curContext.Response.AppendHeader("Content-Disposition",
                "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));

            curContext.Response.BinaryWrite(Export(dtSource, strHeaderText).GetBuffer());
            curContext.Response.End();

        }
        /// <summary>读取excel
        /// 默认第一行为标头
        /// </summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        public static DataTable Import(string strFileName)
        {
            return Import(strFileName, 0);
        }
        /// <summary>
        /// 读取EXCEL
        /// </summary>
        /// <param name="strFileName">EXCEL文件路径</param>
        /// <param name="rowIndex">设置读取行数索引</param>
        /// <returns></returns>
        public static DataTable Import(string strFileName, int rowIndex)
        {
            DataTable dt = new DataTable();
            DataTable ds = new DataTable();
            HSSFWorkbook hssfworkbook;

            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }
            Sheet sheet = hssfworkbook.GetSheetAt(hssfworkbook.ActiveSheetIndex);           
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            Row headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            for (int j = 0; j < cellCount; j++)
            {
                Cell cell = headerRow.GetCell(j);
                if (cell != null)
                {
                    dt.Columns.Add(cell.ToString());
                }

            }

            for (int i = (sheet.FirstRowNum + rowIndex); i <= sheet.LastRowNum; i++)
            {
                Row row = sheet.GetRow(i);
                DataRow dataRow = dt.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        if (!string.IsNullOrEmpty(row.GetCell(j).ToString()))
                        {
                            dataRow[j] = row.GetCell(j).ToString();
                        }
                    }
                }

                dt.Rows.Add(dataRow);
            }
            //清除DATATABLE 空行
            List<DataRow> removelist = new List<DataRow>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool rowdataisnull = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {

                    if (dt.Rows[i][j].ToString().Trim() != "")
                    {

                        rowdataisnull = false;
                    }

                }
                if (rowdataisnull)
                {
                    removelist.Add(dt.Rows[i]);
                }

            }
            for (int i = 0; i < removelist.Count; i++)
            {
                dt.Rows.Remove(removelist[i]);
            }
            return dt;
        }
        public static DataTable Import(Stream ExcelFileStream, string SheetName, int HeaderRowIndex)
        {
            HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileStream);
            Sheet sheet = workbook.GetSheet(SheetName);

            DataTable table = new DataTable();

            Row headerRow = sheet.GetRow(HeaderRowIndex);
            int cellCount = headerRow.LastCellNum;

            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }

            int rowCount = sheet.LastRowNum;

            for (int i = (sheet.FirstRowNum + 1); i < sheet.LastRowNum; i++)
            {
                Row row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();
                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                        dataRow[j] = row.GetCell(j).ToString();
                }
            }
            //清除DATATABLE 空行
            List<DataRow> removelist = new List<DataRow>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                bool rowdataisnull = true;
                for (int j = 0; j < table.Columns.Count; j++)
                {

                    if (table.Rows[i][j].ToString().Trim() != "")
                    {

                        rowdataisnull = false;
                    }

                }
                if (rowdataisnull)
                {
                    removelist.Add(table.Rows[i]);
                }

            }
            for (int i = 0; i < removelist.Count; i++)
            {
                table.Rows.Remove(removelist[i]);
            }
            ExcelFileStream.Close();
            workbook = null;
            sheet = null;
            return table;
        }
        public static DataTable Import(Stream ExcelFileStream, int SheetIndex, int HeaderRowIndex)
        {
                DataTable table = new DataTable();
                HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileStream);
                Sheet sheet = workbook.GetSheetAt(SheetIndex);

                //获取表头
                Row headerRow = sheet.GetRow(HeaderRowIndex);
                
                //string s=headerRow.Sheet.;
                int cellCount = headerRow.LastCellNum;

                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    //string s = headerRow.GetCell(i).StringCellValue;
                    DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                    table.Columns.Add(column);
                }

                int rowCount = sheet.LastRowNum;

                for (int i = (sheet.FirstRowNum + 1); i < sheet.LastRowNum; i++)
                {
                    Row row = sheet.GetRow(i);
                    DataRow dataRow = table.NewRow();
                    if (row != null)
                    {
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                    }
                    table.Rows.Add(dataRow);
                }
                //清除DATATABLE 空行
                List<DataRow> removelist = new List<DataRow>();
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    bool rowdataisnull = true;
                    for (int j = 0; j < table.Columns.Count; j++)
                    {

                        if (table.Rows[i][j].ToString().Trim() != "")
                        {

                            rowdataisnull = false;
                        }

                    }
                    if (rowdataisnull)
                    {
                        removelist.Add(table.Rows[i]);
                    }

                }
                for (int i = 0; i < removelist.Count; i++)
                {
                    table.Rows.Remove(removelist[i]);
                }
                ExcelFileStream.Close();
                workbook = null;
                sheet = null;
            return table;
        }
        #endregion
        #region 利用模版生成EXCEL
        /// <summary>
        /// 利用模板，DataTable导出到Excel（单个类别）
        /// </summary>
        /// <param name="dtSource">DataTable</param>
        /// <param name="strFileName">生成的文件路径、名称</param>
        /// <param name="strTemplateFileName">模板的文件路径、名称</param>
        /// <param name="flg">文件标识（1：个人所得税/2：）</param>
        /// <param name="titleName">表头名称</param>
        public static void ExportExcelForDtByNPOI(DataTable dtSource, string strFileName, string strTemplateFileName, FileEmun.TemplateEmun type, string titleName)
        {
            // 利用模板，DataTable导出到Excel（单个类别）
            using (MemoryStream ms = ExportExcelForDtByNPOI(dtSource, strTemplateFileName, type, titleName))
            {
                byte[] data = ms.ToArray();
                #region 客户端保存
                HttpResponse response = System.Web.HttpContext.Current.Response;
                response.Clear();
                //Encoding pageEncode = Encoding.GetEncoding(PageEncode);
                response.ContentEncoding = Encoding.UTF8;
                response.Charset = "gb2312";
                response.ContentType = "application/vnd.ms-excel";//"application/vnd.ms-excel application/vnd-excel";
                response.AppendHeader("Content-Disposition",
               "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));
                System.Web.HttpContext.Current.Response.BinaryWrite(data);
                #endregion
            }
        }
        /// <summary>
        /// 利用模板，DataTable导出到Excel（单个类别）
        /// </summary>
        /// <param name="dtSource">DataTable</param>
        /// <param name="strTemplateFileName">模板的文件路径、名称</param>
        /// <param name="flg">文件标识--sheet名（1：个人所得税/2：）</param>
        /// <param name="titleName">表头名称</param>
        /// <returns></returns>
        private static MemoryStream ExportExcelForDtByNPOI(DataTable dtSource, string strTemplateFileName, FileEmun.TemplateEmun type, string titleName)
        {

            #region 处理DataTable,处理明细表中没有而需要额外读取汇总值的两列

            #endregion
            int totalIndex = 20;        // 每个类别的总行数
            int rowIndex = 2;       // 起始行

            int dtRowIndex = dtSource.Rows.Count;       // DataTable的数据行数
            FileStream file = new FileStream(strTemplateFileName, FileMode.Open, FileAccess.Read);//读入excel模板
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            string sheetName = "";
            switch (type)
            {
                case FileEmun.TemplateEmun.Template_Income:
                    sheetName = "IncomeReport";
                    break;
                case FileEmun.TemplateEmun.Template_AllBasicPositionSalary:
                    sheetName = "AllBasicSalaryStandardReport";
                    break;
                case FileEmun.TemplateEmun.Template_AreaBasicPositionSalary:
                    sheetName = "AreaBasicPositionSalaryReport";
                    break;
                case FileEmun.TemplateEmun.Template_ShopBasicPositionSalary:
                    sheetName = "ShopBasicPositionSalary";
                    break;
                case FileEmun.TemplateEmun.Template_ExceptionalStaff:
                    sheetName = "ExceptionalStaff";
                    break;
                case FileEmun.TemplateEmun.Template_UserWage:
                    sheetName = "UserWage";
                    break;
            }
            Sheet sheet = workbook.GetSheet(sheetName);

            #region 右击文件 属性信息
            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "易美娜公司";
                workbook.DocumentSummaryInformation = dsi;
                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "于 明"; //填加xls文件作者信息
                si.ApplicationName = "导出EXCEL数据"; //填加xls文件创建程序信息
                si.LastAuthor = "于 明"; //填加xls文件最后保存者信息
                si.Comments = "于 明"; //填加xls文件作者信息
                si.Title = "EXCEL报表数据"; //填加xls文件标题信息
                si.Subject = "EXCEL报表数据";//填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }
            #endregion

            #region 表头
            Row headerRow = sheet.GetRow(0);
            Cell headerCell = headerRow.GetCell(0);
            headerCell.SetCellValue(titleName);
            #endregion

            // 隐藏多余行

            for (int i = rowIndex + dtRowIndex; i < rowIndex + totalIndex; i++)
            {
                Row dataRowD = sheet.GetRow(i);
                dataRowD.Height = 0;
                dataRowD.ZeroHeight = true;
                //sheet.RemoveRow(dataRowD);
            }

            foreach (DataRow row in dtSource.Rows)
            {
                #region 填充内容
                Row dataRow = sheet.GetRow(rowIndex);

                int columnIndex = 0;        // 开始列（0为标题列，从1开始）
                foreach (DataColumn column in dtSource.Columns)
                {
                    // 列序号赋值

                    if (columnIndex > dtSource.Columns.Count)
                        break;

                    Cell newCell = dataRow.GetCell(columnIndex);
                    if (newCell == null)
                        newCell = dataRow.CreateCell(columnIndex);

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String"://字符串类型
                            newCell.SetCellValue(drValue);
                            break;
                        case "System.DateTime"://日期类型
                            DateTime dateV;
                            DateTime.TryParse(drValue, out dateV);
                            newCell.SetCellValue(dateV);

                            break;
                        case "System.Boolean"://布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case "System.Decimal"://浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case "System.DBNull"://空值处理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue("");
                            break;
                    }
                    columnIndex++;
                }
                #endregion

                rowIndex++;
            }
            //格式化当前sheet，用于数据total计算
            sheet.ForceFormulaRecalculation = true;

            #region Clear "0"
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            int cellCount = headerRow.LastCellNum;

            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                Row row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        Cell c = row.GetCell(j);
                        if (c != null)
                        {
                            switch (c.CellType)
                            {
                                case CellType.NUMERIC:
                                    if (c.NumericCellValue == 0)
                                    {
                                        c.SetCellType(CellType.STRING);
                                        c.SetCellValue(string.Empty);
                                    }
                                    break;
                                case CellType.BLANK:

                                case CellType.STRING:
                                    if (c.StringCellValue == "0")
                                    { c.SetCellValue(string.Empty); }
                                    break;

                            }
                        }
                    }
                }
            }
            #endregion

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                sheet = null;
                workbook = null;

                //sheet.Dispose();
                //workbook.Dispose();//一般只用写这一个就OK了，他会遍历并释放所有资源，但当前版本有问题所以只释放sheet
                return ms;
            }
        }
        #endregion
    }
    #endregion
}
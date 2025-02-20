using BrightIdeasSoftware;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XXX
{
    public static class ObjectListViewExtensions
    {
        /// <summary>
        /// 內容匯出Excel
        /// </summary>
        /// <param name="olv"></param>
        public static void ExportToExcel(this ObjectListView olv)
        {
            if (olv.Items.Count == 0)
                return;

            //  初始化 Excel
            Excel.Application excelApp = new Excel.Application { };
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];

            try
            {
                //  將 OLV 的 Column Header 填入 Excel
                for (int i = 0; i < olv.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = olv.AllColumns[i].Text; // Header Text
                }

                //  將 OLV 的資料填入 Excel
                for (int row = 0; row < olv.Items.Count; row++)
                {
                    for (int col = 0; col < olv.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1] = olv.GetModelObject(row)
                                                               .GetType()
                                                               .GetProperty(olv.AllColumns[col].AspectName)?
                                                               .GetValue(olv.GetModelObject(row))?.ToString() ?? "";
                    }
                }

                //  自動調整欄寬
                worksheet.Columns.AutoFit();

                //  顯示 Excel
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error：" + ex.Message);
            }
            finally
            {
                workbook = null;
                worksheet = null;
                excelApp = null;
            }
        }
        /// <summary>
        /// ObjectListView 取得勾選項目轉指定資料模型列舉
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="olv"></param>
        /// <returns></returns>
        public static IEnumerable<T> CheckedList<T>(this ObjectListView olv) where T : class
        {
            return olv.CheckedObjects.Cast<object>().OfType<T>().ToList();
        }
    }
}

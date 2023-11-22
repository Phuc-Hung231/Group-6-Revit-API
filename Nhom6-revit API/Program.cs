using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Reflection;
using System.Drawing;
using System.IO;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
using System.Text.RegularExpressions;
using System.Net;
using Line = Autodesk.Revit.DB.Line;
using System.Xml.Linq;
using static System.Windows.Forms.LinkLabel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;
using Application = Microsoft.Office.Interop.Excel.Application;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Parameter = Autodesk.Revit.DB.Parameter;

namespace Nhom6-revit API
{
    [TransactionAttribute(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ClassMain : IExternalCommand
    {
        UIApplication uiapp;
        UIDocument uidoc;
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        ExternalCommandData revit;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;
            revit = commandData;
            FormMain frm = new FormMain();
            var dlg = frm.ShowDialog();
            if (dlg == System.Windows.Forms.DialogResult.Retry)
            {
                // Lấy các đối tượng đã chọn trong mô hình
                IList<Reference> references = uidoc.Selection.PickObjects(ObjectType.Element, "Chọn một đối tượng");
                List<Element> beams = new List<Element>();
                List<Element> floors = new List<Element>();

                // Phân loại dầm và sàn
                foreach (Reference r in references)
                {
                    Element element = doc.GetElement(r);
                    if (element.Category.Name.Contains("Structural Framing"))
                    {
                        beams.Add(element);
                    }
                    else if (element.Category.Name.Contains("Floors"))
                    {
                        floors.Add(element);
                    }
                }
                
                frm.kichthuoc(beams[0], floors[0]);
                dlg = frm.ShowDialog();

            }
            if (dlg == System.Windows.Forms.DialogResult.Ignore)
            {
                string filePath = "";
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Title = "Chọn tệp";
                    openFileDialog.Filter = "Tất cả tệp|*.*";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        
                        filePath = openFileDialog.FileName;
                    }
                }
                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(filePath);
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1]; // Lấy trang tính đầu tiên

                Range range = worksheet.UsedRange;
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    List<string> lines = new List<string>();
                    for (int j = 1; j <= colCount; j++)
                    {
                        try
                        {
                            var cellValue = (range.Cells[i, j] as Range).Value2;
                            lines.Add(cellValue.ToString());

                        }
                        catch { }

                    }
                    //frm.importable(lines);
                }
                // Đóng Excel
                workbook.Close(false);
                excelApp.Quit();

                // Giải phóng tài nguyên
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(excelApp);
                dlg = frm.ShowDialog();

            }
            return Result.Succeeded;

        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }

}
public class ColumnSelectionFilter : ISelectionFilter
{
    public bool AllowElement(Element elem)
    {
        // Kiểm tra xem phần tử có phải là cột không
        if (elem.Category != null && elem.Category.Name == "Columns")
        {
            return true;
        }
        return false;
    }

    public bool AllowReference(Reference reference, XYZ position)
    {
        return false;
    }
}


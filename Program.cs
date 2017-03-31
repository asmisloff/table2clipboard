using System.Text;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;

namespace tbl2clibpoard
{
    public class Program
    {
        [CommandMethod("ttcb")]
        public void ttcb() {
            IAcadApplication acad = System.Runtime.InteropServices.Marshal.GetActiveObject("Autocad.Application") as IAcadApplication;
            if (acad != null) {
                var doc = acad.ActiveDocument;
                AcadSelectionSet sset = null;
                foreach (AcadSelectionSet ss in doc.SelectionSets) {
                    if (ss.Name == "tbl2clipboard_sset") {
                        sset = ss;
                    }
                }
                if (sset == null) {
                    sset = doc.SelectionSets.Add("tbl2clipboard_sset");
                }
                sset.SelectOnScreen();
                foreach (AcadEntity ent in sset) {
                    if (ent.ObjectName == "AcDbTable") {
                        AcadTable t = ent as AcadTable;
                        StringBuilder table = new StringBuilder(t.Rows - 1);
                        for (int r = 1; r < t.Rows; r++) {
                            string row = "";
                            for (int c = 0; c < t.Columns; c++) {
                                string cv = t.GetCellValue(r, c);
                                if (c == 2) {
                                    cv = cv.Replace('x', '\t');
                                }
                                row += cv + '\t';
                            }
                            table.Append(row.TrimEnd('\t') + '\n');
                        }
                        Clipboard.SetText(table.ToString());
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("ok");
                    }
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace ExcelAddin
{
    public interface IExcelAddinAndrei
    {
        int MedianRange(Microsoft.Office.Interop.Excel.Range r);
    }

    [Guid("8E869C7B-B654-4E55-9838-3989CBCB6AD1"), ProgId("ExcelAddinAndrei.Median"), ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class ExcelAddinAndreiTest : IExcelAddinAndrei
    {
        
        public int MedianRange(Microsoft.Office.Interop.Excel.Range r)
        {
            List<int> data = new List<int>();

            foreach (Microsoft.Office.Interop.Excel.Range row in r.Rows)
            {
                foreach (Microsoft.Office.Interop.Excel.Range cell in row.Columns)
                {
                    data.Add((int)cell.Value);
                }
            }

            data.Sort();

            if (data.Capacity % 2 == 0)
            {
                int index1 = data.Capacity / 2 - 1;
                int index2 = index1 + 1;

                return (data[index1] + data[index2]) / 2;
            }
            else
            {
                return data[data.Capacity / 2];
            }
        }

    }
}

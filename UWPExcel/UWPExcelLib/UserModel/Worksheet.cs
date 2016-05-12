// Created by Thiry Patrick 12/05/2016

using System.Collections.Generic;

namespace UWPExcelLib.UserModel
{
    public class Worksheet
    {
        public string Name;

        public List<Row> Rows;

        public Worksheet(string NameSheet, List<Row> rows)
        {
            Name = NameSheet;
            Rows = rows;
        }

        public Row GetRow(int NumRow)
        {
            foreach (Row r in Rows)
            {
                if (r.NumRow == NumRow)
                {
                    return r;
                }
            }
            return null;
        }

        public Cell GetCell(int NumRow, string Letter)
        {
            foreach (Row r in Rows)
            {
                if (r.NumRow == NumRow)
                {
                    foreach (Cell c in r.Cells)
                    {
                        if (c.NumCol == Letter)
                            return c;
                    }
                }
            }
            return null;
        }
    }
}

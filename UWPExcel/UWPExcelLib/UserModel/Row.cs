// Created by Thiry Patrick 12/05/2016

using System.Collections.Generic;

namespace UWPExcelLib.UserModel
{
    public class Row
    {
        public int NumRow;
        public List<Cell> Cells;

        // Constructeur en private
        private Row() { }

        // Constructeur principal
        public Row(int numRow, List<Cell> cells)
        {
            NumRow = numRow;
            Cells = cells;
        }
    }
}

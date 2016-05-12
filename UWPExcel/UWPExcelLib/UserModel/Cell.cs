// Created by Thiry Patrick 12/05/2016


namespace UWPExcelLib.UserModel
{
    public class Cell
    {
        private string _numCol;
        public string NumCol
        {
            private set { _numCol = value; }
            get { return _numCol; }
        }

        private string _value;
        public string Value
        {
            private set { _value = value; }
            get { return _value; }
        }

        // Constructeur privé
        private Cell() { }

        // Constructeur publique
        public Cell(string numCol, string value)
        {
            NumCol = numCol;
            Value = value;
        }
    }
}

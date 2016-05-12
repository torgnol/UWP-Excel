// Created by Thiry Patrick 12/05/2016

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Windows.Storage;

namespace UWPExcelLib.UserModel
{
    public class Workbook
    {
        private const string BASE_FILE = "[Content_Types].xml";

        // ContentType
        private const string CONTENT_SHAREDSTRING = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        private const string CONTENT_STYLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        private const string CONTENT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";
        private const string CONTENT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string CONTENT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

        /// <summary>
        /// XLSX File location
        /// </summary>
        private StorageFile file;

        /// <summary>
        /// Internal zip file
        /// </summary>
        private string sharedStringFile;
        private string styleFile;
        private string themeFile;
        private string workbookFile;
        private List<string> listWorksheetFile;

        /// <summary>
        /// DataFile
        /// </summary>
        private List<string> SharedString;
        private List<string> WorksheetName;

        /// <summary>
        /// OpenFile
        /// </summary>
        /// <param name="file"></param>
        public void Open(StorageFile file)
        {
            this.file = file;

            ZipArchive z = new ZipArchive(file.OpenStreamForReadAsync().Result);
            var contentFile = z.GetEntry(BASE_FILE);

            // Liste des worksheets
            listWorksheetFile = new List<string>();

            // OpenFile - List file
            using (var sr = contentFile.Open())
            {
                XDocument xdoc = XDocument.Load(sr);

                foreach (XElement e in xdoc.Root.Elements())
                {
                    switch (e.Attribute("ContentType").Value)
                    {
                        case CONTENT_SHAREDSTRING:
                            sharedStringFile = e.Attribute("PartName").Value;
                            sharedStringFile = sharedStringFile.Remove(0, 1);
                            break;
                        case CONTENT_STYLE:
                            styleFile = e.Attribute("PartName").Value;
                            styleFile = styleFile.Remove(0, 1);
                            break;
                        case CONTENT_THEME:
                            themeFile = e.Attribute("PartName").Value;
                            themeFile = themeFile.Remove(0, 1);
                            break;
                        case CONTENT_WORKBOOK:
                            workbookFile = e.Attribute("PartName").Value;
                            workbookFile = workbookFile.Remove(0, 1);
                            break;
                        case CONTENT_WORKSHEET:
                            string worksheetFile = e.Attribute("PartName").Value;
                            worksheetFile = worksheetFile.Remove(0, 1);
                            listWorksheetFile.Add(worksheetFile);
                            break;
                    }
                }
            }

            // SharedString
            if (sharedStringFile != null)
            {
                var sharedString = z.GetEntry(sharedStringFile);
                using (var sr = sharedString.Open())
                {
                    XDocument xdoc = XDocument.Load(sr);
                    SharedString =
                        (
                        from e in xdoc.Root.Elements()
                        select e.Elements().First().Value
                        ).ToList();
                }
            }

            // Sheet Names
            WorksheetName = new List<string>();
            var worksheetName = z.GetEntry(workbookFile);
            using (var wk = worksheetName.Open())
            {
                XDocument xdoc = XDocument.Load(wk);
                XNamespace xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                XElement sheetData = xdoc.Root.Element(xmlns + "sheets");

                foreach (XElement e in sheetData.Elements())
                {
                    string name = e.Attribute("name").Value;
                    WorksheetName.Add(name);
                }
            }
        }

        /// <summary>
        /// Return sheet count
        /// </summary>
        /// <returns></returns>
        public int GetSheetCount()
        {
            return listWorksheetFile.Count;
        }

        /// <summary>
        /// Sheet by Id (list)
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Worksheet GetSheetAt(int id)
        {
            ZipArchive z = new ZipArchive(file.OpenStreamForReadAsync().Result);
            var worksheetFile = z.GetEntry(listWorksheetFile[id]);

            // Récupération de l'ensemble de la feuille
            List<Row> rows = new List<Row>();
            using (var wk = worksheetFile.Open())
            {
                XDocument xdoc = XDocument.Load(wk);
                XNamespace xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                XElement sheetData = xdoc.Root.Element(xmlns + "sheetData");
                // foreach row
                foreach (XElement e in sheetData.Elements())
                {
                    // ligne
                    string numero = e.Attribute("r").Value;

                    // Init column
                    List<Cell> columns = new List<Cell>();

                    // foreach column
                    foreach (XElement f in e.Elements())
                    {
                        string colonne = f.Attribute("r").Value;
                        colonne = colonne.Replace(numero, "");

                        // Data Treatment
                        if (f.Attribute("t") != null)
                            columns.Add(new Cell(colonne, SharedString[int.Parse(f.Value)]));
                        else
                            columns.Add(new Cell(colonne, f.Value));
                    }

                    // Enregistrement de la ligne
                    rows.Add(new Row(int.Parse(numero), columns));
                }
            }

            return new Worksheet(WorksheetName[id], rows);
        }

        /// <summary>
        /// Sheet by Id (list)
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public Worksheet GetSheetAt(string name)
        {
            int count = 0;
            foreach (string _name in WorksheetName)
            {
                if (_name == name)
                {
                    return this.GetSheetAt(count);
                }
                count++;
            }
            return null;
        }
    }
}
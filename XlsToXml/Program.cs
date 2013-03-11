using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Xml;

namespace XlsToXml
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var f = @"c:\xls.xls";
            int page = 0;
            var o = "c:\\xls.xml";
            var type = "csv";
            List<string> names = new List<string>();
            if (args.Length > 0)
            {
                if (args[0].Trim().ToLower() == "help")
                {
                    Console.WriteLine("** HELP **");
                    Console.WriteLine("argument 0: xml,csv");
                    Console.WriteLine("argument 1: input filename");
                    Console.WriteLine("argument 2: page number to process");
                    Console.WriteLine("argument 3: output filename");
                    Console.WriteLine("argument 4 to x: param names in xml file only");
                    return;

                }
                f = args[1];
                page = int.Parse(args[2]);
                o = args[3];
                for (int x = 4; x < args.Length; x++)
                {
                    names.Add(args[x]);
                }
            }


            String sConnectionString =
                "Provider=Microsoft.Jet.OLEDB.4.0;" +
                "Data Source=" + f + ";" +
                "Extended Properties=Excel 8.0;";


            OleDbConnection objConn = new OleDbConnection(sConnectionString);

            objConn.Open();

            var dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            String[] excelSheets = new String[dt.Rows.Count];
            int i = 0;

// Add the sheet name to the string array.
            foreach (DataRow row in dt.Rows)
            {
                excelSheets[i] = row["TABLE_NAME"].ToString();
                i++;
            }


            OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [" + excelSheets[page] + "]", objConn);

            OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();

            objAdapter1.SelectCommand = objCmdSelect;

            DataSet ds = new DataSet();

            objAdapter1.Fill(ds);

            objConn.Close();

            if (type == "xml")
            {

                XmlDocument doc = new XmlDocument(); // Create the XML Declaration, and append it to XML document
                XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", null, null);
                doc.AppendChild(dec); // Create the root element
                XmlElement root = doc.CreateElement("root");
                doc.AppendChild(root);

                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    var c = row.ItemArray.Count();
                    if (names.Count > 0)
                    {
                        c = names.Count;
                    }
                    XmlElement r = doc.CreateElement("row");
                    for (var n = 0; n < c; n++)
                    {
                        var itemName = "item";
                        if (names.Count > 0)
                        {
                            itemName = names[n];
                        }
                        XmlElement item = doc.CreateElement(itemName);
                        XmlCDataSection idata = doc.CreateCDataSection(row.ItemArray[n].ToString());
                        item.AppendChild(idata);
                        r.AppendChild(item);
                    }
                    root.AppendChild(r);
                }

                string xmlOutput = doc.OuterXml;


                // Write the string to a file.
                System.IO.StreamWriter file = new System.IO.StreamWriter(o);
                file.Write(xmlOutput);
                file.Close();
            }else if (type == "csv")
            {
                var sb = new StringBuilder();
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    var c = row.ItemArray.Count();
                    if (names.Count > 0)
                    {
                        c = names.Count;
                    }
                    for (var n = 0; n < c; n++)
                    {
                        var itemName = "item"; //name of param
                        var itemValue = row.ItemArray[n].ToString();
                        sb.Append(itemValue + ",");
                    }
                    sb.Remove(sb.Length - 1, 1);
                    sb.Append(Environment.NewLine);
                }



                // Write the string to a file.
                System.IO.StreamWriter file = new System.IO.StreamWriter(o);
                file.Write(sb.ToString());
                file.Close();
            }
        }
    }
}
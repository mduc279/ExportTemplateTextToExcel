using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Xml;

namespace ExportTemplateTextToExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var filePath = "C:\\Users\\DucDoanMinh\\OneDrive - Add-On Products & Add-On Development\\Desktop\\Colorful_Sokolow Ed. 1_TestPolishLanguage v_4_3.0 (1280x800) English.xml";
            var destinationLanguage = "Polish";
            try
            {
                var allText = File.ReadAllText(filePath);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(allText);
                var textNodes = doc.SelectNodes("//Text");
                var listTextNodes = textNodes.Cast<XmlNode>().Select(e => e.InnerText).Distinct().ToList();

                var placeholderNodes = doc.SelectNodes("//Placeholder");
                var listPlaceholder = placeholderNodes.Cast<XmlNode>().Select(e => e.InnerText).Distinct().ToList();

                var messageInfoNodes = doc.SelectNodes("//MessageInfo/*");
                var listMessageInfo = messageInfoNodes.Cast<XmlNode>().Select(e => e.InnerText).Distinct().ToList();

                var bookName = doc.SelectSingleNode("//TemplateInfo/Name").InnerText;
                var destination = $"C:\\Users\\DucDoanMinh\\OneDrive - Add-On Products & Add-On Development\\Desktop\\{bookName}.xlsx";

                var workBook = new XLWorkbook();
                workBook.AddWorksheet("Text");
                workBook.AddWorksheet("Placeholder");
                workBook.AddWorksheet("MessageInfo");

                int row = 1;
                var txtSheet = workBook.Worksheet("Text");
                txtSheet.Cell("A" + row.ToString()).Value = "English";
                txtSheet.Cell("B" + row.ToString()).Value = destinationLanguage;
                foreach (var item in listTextNodes)
                {
                    row++;
                    txtSheet.Cell("A" + row.ToString()).Value = item;
                }

                row = 1;
                var placeholderSheet = workBook.Worksheet("Placeholder");
                placeholderSheet.Cell("A" + row.ToString()).Value = "English";
                placeholderSheet.Cell("B" + row.ToString()).Value = destinationLanguage;
                foreach (var item in listPlaceholder)
                {
                    row++;
                    placeholderSheet.Cell("A" + row.ToString()).Value = item;
                }

                row = 1;
                var messageInfoSheet = workBook.Worksheet("MessageInfo");
                messageInfoSheet.Cell("A" + row.ToString()).Value = "English";
                messageInfoSheet.Cell("B" + row.ToString()).Value = destinationLanguage;
                foreach (var item in listMessageInfo)
                {
                    row++;
                    messageInfoSheet.Cell("A" + row.ToString()).Value = item;
                }
                
                workBook.SaveAs(destination);
            }
            catch (Exception ex)
            {

            }
        }
    }
}

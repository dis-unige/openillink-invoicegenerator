using ClosedXML.Excel;
using FileHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using Openillink.InvoiceGenerator.Models;

namespace Openillink.InvoiceGenerator
{
    public class InvoiceGenerator
    {
        public void Run(Options options)
        {
            try
            {
                Console.WriteLine($"Processing orders in '{options.InputFile}' from {options.StartDate:d} to {options.EndDate:d}...");
                var startDate = options.StartDate;
                var endDate = options.EndDate;
                var data = ReadData(options.InputFile, startDate, endDate);

                foreach (var group in data.GroupBy(d => d.InvoiceGroup))
                {
                    var filename = string.Format("{0}_{1:yyyyMMdd}_{2:yyyyMMdd}.xlsx", group.Key, startDate, endDate);
                    CreateWorkbook(filename, group.Key, group.ToList(), startDate, endDate);
                }

                Console.WriteLine("Done...");
            }
            catch (Exception ex)
            {
                var color = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("An error occured!");
                Console.WriteLine(ex.Message);
                Console.ForegroundColor = color;
            }
        }

        private static IList<Order> ReadData(string file, DateTime startDate, DateTime endDate)
        {
            var engine = new FileHelperEngine<Order>();
            var records = engine.ReadFile(file)
                                .Where(r => r.SendDate >= startDate && r.SendDate <= endDate).ToList();

            // var replaceByArve = new[] { "BFS", "Anthropologie", "CUI", "Math", "ISE" };

            // foreach (var item in records)
            // {
            // if (item.Localisation == "BFM")
            // {
            // item.Localisation = "CMU";
            // }
            // else if (replaceByArve.Contains(item.Localisation))
            // {
            // item.Localisation = "Arve";
            // }
            // }

            return records;
        }

        private static void CreateWorkbook(string filename, string key, IList<Order> orders, DateTime startDate, DateTime endDate)
        {
            Console.WriteLine("Creating file {0}, Count: {1}", filename, orders.Count);

            var wb = new XLWorkbook();

            AddOrdersWorksheet(wb.AddWorksheet(key), orders);

            var groups = orders.GroupBy(o => o.InvoiceAccount);

            foreach (var item in groups.Where(g => !string.IsNullOrEmpty(g.Key)).OrderBy(g => g.Key))
            {
                AddDetailsWorksheet(wb.AddWorksheet(item.Key), item.ToList(), startDate, endDate);
            }

            AddSummaryWorksheet(wb.AddWorksheet("Controle", 2), groups.Where(g => !string.IsNullOrEmpty(g.Key)).OrderBy(g => g.Key).Select(g => g.Key).ToList());

            wb.SaveAs(filename);
        }

        private static void AddSummaryWorksheet(IXLWorksheet sheet, List<string> list)
        {
            sheet.Row(1).Cell(1).Value = "Compte";
            sheet.Row(1).Cell(2).Value = "Nombre";
            sheet.Row(1).Cell(3).Value = "Prix";

            var index = 2;

            foreach (var item in list)
            {
                sheet.Row(index).Cell(1).Value = item;
                sheet.Row(index).Cell(2).FormulaA1 = $"={item}!_{item}_COUNT";
                sheet.Row(index).Cell(3).FormulaA1 = $"={item}!_{item}_PRICE";
                sheet.Row(index).Cell(3).Style.NumberFormat.Format = "#,##0.00";

                index++;
            }

            var used = sheet.RangeUsed().Rows(1, index - 1);
            used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        private static void AddOrdersWorksheet(IXLWorksheet sheet, IList<Order> orders)
        {
            var cell = 1;
            sheet.Row(1).Cell(cell++).Value = "Compte";
            sheet.Row(1).Cell(cell++).Value = "Nom";
            sheet.Row(1).Cell(cell++).Value = "Prénom";
            sheet.Row(1).Cell(cell++).Value = "Email";
            sheet.Row(1).Cell(cell++).Value = "N° commande";
            sheet.Row(1).Cell(cell++).Value = "Date commande";
            sheet.Row(1).Cell(cell++).Value = "Date envoi";
            sheet.Row(1).Cell(cell++).Value = "Prix";
            sheet.Row(1).Cell(cell++).Value = "Localisation";
            sheet.Row(1).Cell(cell++).Value = "Type Doc";
            sheet.Row(1).Cell(cell++).Value = "Périodique";
            sheet.Row(1).Cell(cell++).Value = "Année";
            sheet.Row(1).Cell(cell++).Value = "Vol";
            sheet.Row(1).Cell(cell++).Value = "pp.";

            var index = 2;
            sheet.ColumnWidth = 11;

            foreach (var order in orders)
            {
                sheet.Row(index).Cell(1).Value = order.InvoiceAccount;
                sheet.Row(index).Cell(2).Value = order.Name;
                sheet.Row(index).Cell(3).Value = order.FirstName;
                sheet.Row(index).Cell(4).Value = order.EMail;

                sheet.Row(index).Cell(5).Value = order.IlLinkId;

                sheet.Row(index).Cell(6).Style.DateFormat.Format = "dd.MM.yyyy";
                sheet.Row(index).Cell(6).Value = order.OrderDate;

                sheet.Row(index).Cell(7).Style.DateFormat.Format = "dd.MM.yyyy";
                sheet.Row(index).Cell(7).Value = order.SendDate;

                sheet.Row(index).Cell(8).Style.NumberFormat.Format = "#,##0.00";
                sheet.Row(index).Cell(8).Value = order.Price;

                sheet.Row(index).Cell(9).Value = order.Localisation;
                sheet.Row(index).Cell(10).Value = order.DocumentType;
                sheet.Row(index).Cell(11).Value = order.Title;
                sheet.Row(index).Cell(12).Value = order.Year;
                sheet.Row(index).Cell(13).Value = order.Volume;
                sheet.Row(index).Cell(14).Value = order.Pages;

                index++;
            }

            var tablename = $"Table_{sheet.Name}";
            var table = sheet.Range(sheet.FirstColumn().FirstCell(), sheet.LastRowUsed().LastCellUsed()).CreateTable(tablename);
            table.ShowAutoFilter = true;
            table.Theme = XLTableTheme.TableStyleMedium2;
        }

        private static void AddDetailsWorksheet(IXLWorksheet sheet, IList<Order> orders, DateTime startDate, DateTime endDate)
        {
            sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;

            var cell = 1;
            sheet.Row(1).Cell(cell++).Value = "Compte";
            sheet.Row(1).Cell(cell++).Value = "Nom";
            sheet.Row(1).Cell(cell++).Value = "Prénom";
            sheet.Row(1).Cell(cell++).Value = "Email";
            sheet.Row(1).Cell(cell++).Value = "N° commande";
            sheet.Row(1).Cell(cell++).Value = "Date envoi";
            sheet.Row(1).Cell(cell++).Value = "Prix";
            sheet.Row(1).Cell(cell++).Value = "Localisation";
            sheet.Row(1).Cell(cell++).Value = "Type Doc";
            sheet.Row(1).Cell(cell++).Value = "Périodique";
            sheet.Row(1).Cell(cell++).Value = "Année";
            sheet.Row(1).Cell(cell++).Value = "Vol";
            sheet.Row(1).Cell(cell++).Value = "pp.";

            var index = 2;
            sheet.ColumnWidth = 11;

            foreach (var order in orders)
            {
                sheet.Row(index).Cell(1).Value = order.InvoiceAccount;
                sheet.Row(index).Cell(2).Value = order.Name;
                sheet.Row(index).Cell(3).Value = order.FirstName;
                sheet.Row(index).Cell(4).Value = order.EMail;

                sheet.Row(index).Cell(5).Value = order.IlLinkId;

                sheet.Row(index).Cell(6).Style.DateFormat.Format = "dd.MM.yyyy";
                sheet.Row(index).Cell(6).Value = order.SendDate;

                sheet.Row(index).Cell(7).Style.NumberFormat.Format = "#,##0.00";
                sheet.Row(index).Cell(7).Value = order.Price;

                sheet.Row(index).Cell(8).Value = order.Localisation;
                sheet.Row(index).Cell(9).Value = order.DocumentType;
                sheet.Row(index).Cell(10).Value = order.Title;
                sheet.Row(index).Cell(11).Value = order.Year;
                sheet.Row(index).Cell(12).Value = order.Volume;
                sheet.Row(index).Cell(13).Value = order.Pages;

                index++;
            }

            var lastDataRow = index - 1;

            var tablename = $"Table_{sheet.Name}";
            var table = sheet.Range(1, 1, index - 1, 13).CreateTable(tablename);
            table.ShowAutoFilter = false;
            table.Theme = XLTableTheme.TableStyleMedium2;

            sheet.Row(index).Cell(1).FormulaA1 = $"=COUNT(G2:G{lastDataRow})";
            sheet.Row(index).Cell(1).AddToNamed("Count");

            sheet.Row(index).Cell(2).Value = string.Format("commandes du {0:dd MMMM} au {1:dd MMMM yyyy}", startDate, endDate);
            sheet.Row(index).Cell(6).Value = "Total:";

            sheet.Row(index).Cell(7).Style.NumberFormat.Format = "#,##0.00";
            sheet.Row(index).Cell(7).FormulaA1 = $"=SUM(G2:G{lastDataRow})";

            sheet.Range(index, 1, index, 1).AddToNamed($"_{sheet.Name}_COUNT", XLScope.Worksheet);
            sheet.Range(index, 7, index, 7).AddToNamed($"_{sheet.Name}_PRICE", XLScope.Worksheet);

            sheet.LastRowUsed().Cells(1, 13).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            sheet.LastRowUsed().Cells(1, 13).Style.Font.Bold = true;

            // Add table if more than 10 order or more than 1 location
            if (orders.Count >= 10 || orders.Select(o => o.Localisation).Distinct().Count() != 1)
            {
                AddSummaryTable(sheet, index, lastDataRow);
            }
        }

        private static void AddSummaryTable(IXLWorksheet sheet, int index, int lastDataRow)
        {
            index += 2;

            sheet.Row(index).Cell(3).FormulaR1C1 = string.Format("=COUNTIF(H:H,\"CMU\")+COUNTIF(H:H,\"CMU_AZ\")+COUNTIF(H:H,\"CMU_UNIGE\")+COUNTIF(H:H,\"CMU_DBU\")+COUNTIF(H:H,\"CMU_HUG\")+COUNTIF(H:H,\"CMU_IEH2\")+COUNTIF(H:H,\"ARVE-UNIGE\")+COUNTIF(H:H,\"ARVE-BELS\")+COUNTIF(H:H,\"ARVE-CUI\")+COUNTIF(H:H,\"ARVE-DBU\")+COUNTIF(H:H,\"ARVE-ISE\")+COUNTIF(H:H,\"ARVE_AZ\")+COUNTIF(H:H,\"ARVE-MATH\")+COUNTIF(H:H,\"ARVE-OBS\")+COUNTIF(H:H,\"ARVE-TERRE\")+COUNTIF(H:H,\"ARVE\")", lastDataRow);
            sheet.Row(index).Cell(4).Value = "commandes UNIGE";
            sheet.Row(index).Cell(7).Style.NumberFormat.Format = "#,##0.00";
            sheet.Row(index).Cell(7).FormulaR1C1 = string.Format("=SUMIF(H:H,\"CMU\",G:G)+SUMIF(H:H,\"CMU_AZ\",G:G)+SUMIF(H:H,\"CMU_UNIGE\",G:G)+SUMIF(H:H,\"CMU_DBU\",G:G)+SUMIF(H:H,\"CMU_HUG\",G:G)+SUMIF(H:H,\"CMU_IEH2\",G:G)+SUMIF(H:H,\"ARVE-UNIGE\",G:G)+SUMIF(H:H,\"ARVE-BELS\",G:G)+SUMIF(H:H,\"ARVE-CUI\",G:G)+SUMIF(H:H,\"ARVE-DBU\",G:G)+SUMIF(H:H,\"ARVE-ISE\",G:G)+SUMIF(H:H,\"ARVE_AZ\",G:G)+SUMIF(H:H,\"ARVE-MATH\",G:G)+SUMIF(H:H,\"ARVE-OBS\",G:G)+SUMIF(H:H,\"ARVE-TERRE\",G:G)+SUMIF(H:H,\"ARVE\",G:G)", lastDataRow);

            index++;

            sheet.Row(index).Cell(3).FormulaR1C1 = string.Format("=COUNTIF(H:H,\"CMU_GE\")+COUNTIF(H:H,\"CMU_AUTRESUISSE\")+COUNTIF(H:H,\"CMU_IDS\")+COUNTIF(H:H,\"CMU_ILLRERO\")+COUNTIF(H:H,\"CMU_NEBIS\")+COUNTIF(H:H,\"CMU_RENOUVAUD\")+COUNTIF(H:H,\"CMU_NEBIS\")+COUNTIF(H:H,\"ARVE-GE\")+COUNTIF(H:H,\"ARVE_AUTRESUISSE\")+COUNTIF(H:H,\"ARVE-IDS\")+COUNTIF(H:H,\"ARVE-ILLRERO\")+COUNTIF(H:H,\"ARVE-NEBIS\")+COUNTIF(H:H,\"ARVE-RENOUVAUD\")", lastDataRow);
            sheet.Row(index).Cell(4).Value = "commandes Suisse";
            sheet.Row(index).Cell(7).Style.NumberFormat.Format = "#,##0.00";
            sheet.Row(index).Cell(7).FormulaR1C1 = string.Format("=SUMIF(H:H,\"CMU_GE\",G:G)+SUMIF(H:H,\"CMU_AUTRESUISSE\",G:G)+SUMIF(H:H,\"CMU_IDS\",G:G)+SUMIF(H:H,\"CMU_ILLRERO\",G:G)+SUMIF(H:H,\"CMU_NEBIS\",G:G)+SUMIF(H:H,\"CMU_RENOUVAUD\",G:G)+SUMIF(H:H,\"CMU_NEBIS\",G:G)+SUMIF(H:H,\"ARVE-GE\",G:G)+SUMIF(H:H,\"ARVE_AUTRESUISSE\",G:G)+SUMIF(H:H,\"ARVE-IDS\",G:G)+SUMIF(H:H,\"ARVE-ILLRERO\",G:G)+SUMIF(H:H,\"ARVE-NEBIS\",G:G)+SUMIF(H:H,\"ARVE-RENOUVAUD\",G:G)", lastDataRow);

            index++;

            // sheet.Row(index).Cell(3).FormulaR1C1 = string.Format("=COUNTIF(H:H,\"SUBITO\")", lastDataRow);
            // sheet.Row(index).Cell(4).Value = "commandes Subito (All.)";
            // sheet.Row(index).Cell(7).Style.NumberFormat.Format = "#,##0.00";
            // sheet.Row(index).Cell(7).FormulaR1C1 = string.Format("=SUMIF(H:H,\"SUBITO\",G:G)", lastDataRow);
            // 
            // index++;

            sheet.Row(index).Cell(3).FormulaR1C1 = string.Format("=COUNTIF(H:H,\"CMU_ETRANGER\")+COUNTIF(H:H,\"CMU_BL\")+COUNTIF(H:H,\"CMU_NLM\")+COUNTIF(H:H,\"CMU_SUBITO\")+COUNTIF(H:H,\"CMU_SUDOC\")+COUNTIF(H:H,\"ARVE-ETRANGER\")+COUNTIF(H:H,\"ARVE-NLM\")+COUNTIF(H:H,\"ARVE-SUBITO\")+COUNTIF(H:H,\"ARVE-SUDOC\")", lastDataRow);
            sheet.Row(index).Cell(4).Value = "commandes Etranger";
            sheet.Row(index).Cell(7).Style.NumberFormat.Format = "#,##0.00";
            sheet.Row(index).Cell(7).FormulaR1C1 = string.Format("=SUMIF(H:H,\"CMU_ETRANGER\",G:G)+SUMIF(H:H,\"CMU_BL\",G:G)+SUMIF(H:H,\"CMU_NLM\",G:G)+SUMIF(H:H,\"CMU_SUBITO\",G:G)+SUMIF(H:H,\"CMU_SUDOC\",G:G)+SUMIF(H:H,\"ARVE-ETRANGER\",G:G)+SUMIF(H:H,\"ARVE-NLM\",G:G)+SUMIF(H:H,\"ARVE-SUBITO\",G:G)+SUMIF(H:H,\"ARVE-SUDOC\",G:G)", lastDataRow);

            index++;

            sheet.Row(index).Cell(3).FormulaR1C1 = string.Format("=COUNTIF(H:H,\"CMU_OA\")+COUNTIF(H:H,\"ARVE_OA\")", lastDataRow);
            sheet.Row(index).Cell(4).Value = "Open Access online";
            sheet.Row(index).Cell(7).Style.NumberFormat.Format = "#,##0.00";
            sheet.Row(index).Cell(7).FormulaR1C1 = string.Format("=SUMIF(H:H,\"CMU_OA\",G:G)+SUMIF(H:H,\"ARVE_OA\",G:G)", lastDataRow);

            sheet.Row(index).Cells(3, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thick;

            index++;

            sheet.Row(index).Cells(3, 7).Style.Font.Bold = true;
            sheet.Row(index).Cell(3).FormulaA1 = string.Format("=SUM(C{0}:C{1})", index - 5, index - 1);
            sheet.Row(index).Cell(7).Style.NumberFormat.Format = "#,##0.00";
            sheet.Row(index).Cell(7).FormulaA1 = string.Format("=SUM(G{0}:G{1})", index - 5, index - 1);
        }
    }
}

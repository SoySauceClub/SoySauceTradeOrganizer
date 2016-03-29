using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using System.IO;

namespace SoySauceTradeOrganizer
{
    class Program
    {
        public const string ENTRY_FILE = @"TradeCharts\{0}_{1}_Entry.png";
        public const string EXIT_FILE = @"TradeCharts\{0}_{1}_Exit.png";
        public const string EXIT_PAST_10 = @"TradeCharts\{0}_{1}_Exit_p10.png";
        public const string EXIT_PAST_30 = @"TradeCharts\{0}_{1}_Exit_p30.png";
        static void Main(string[] args)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            //var ef = ExcelFile.Load(@"I:\Dropbox\Daniel\TestTrades\TripleScreenBull\bull_trade_result.xlsx");

            var reader = new StreamReader(File.OpenRead(@"C:\temp\TradeConfig\bull_trade_result - Copy.csv"));

            var tradeList = new List<TradeResult>();

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (line.StartsWith("EnterDate"))
                    continue;

                var values = line.Split(',');

                var tradeResult = new TradeResult
                {
                    EntryDate = DateTime.Parse(values[0]),
                    EnterDirection = values[1],
                    EnterPrice = decimal.Parse(values[2]),
                    Ticker = values[3],
                    ExitDate = DateTime.Parse(values[4]),
                    ExitPrice = decimal.Parse(values[5]),
                    InitialStop = decimal.Parse(values[6]),
                    isMissed = int.Parse(values[7]),
                    TargetPrice = decimal.Parse(values[8])
                };

                tradeList.Add(tradeResult);
            }

            var ef = new ExcelFile();
            var ws = ef.Worksheets.Add("TradeResult");
            var excel_row_counter = 1;
            for(var i = 0; i < tradeList.Count; i++)
            {
                // the free version of GemBox Excel only allow 150 lines, start a new file if reach to 150 rows
                if (i % 150 == 0 && i != 0)
                {
                    ef.Save(string.Format("tradeResult{0}.xlsx", i.ToString()));
                    ef = new ExcelFile();
                    ws = ef.Worksheets.Add("TradeResult");
                    excel_row_counter = 1;
                }
                var tradeResult = tradeList[i];
                ws.Cells["A" + excel_row_counter.ToString()].Value = tradeResult.Ticker;
                ws.Cells["B" + excel_row_counter.ToString()].Value = tradeResult.EnterDirection;

                var entryDateCell = ws.Cells["C" + excel_row_counter.ToString()];
                entryDateCell.Value = tradeResult.EntryDate;
                entryDateCell.Style.NumberFormat = "MM/dd/yyyy";


                var cell = ws.Cells["D" + excel_row_counter.ToString()];
                var entry_link = string.Format(ENTRY_FILE, tradeResult.EntryDate.ToString("dd/MM/yyyy"), tradeResult.Ticker);
                cell.Formula = string.Format("=HYPERLINK(\"{0}\", \"{1}\")", entry_link, tradeResult.EnterPrice);
                cell.Style.Font.UnderlineStyle = UnderlineStyle.Single;
                cell.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);

                ws.Cells["E" + excel_row_counter.ToString()].Value = tradeResult.InitialStop;
                ws.Cells["F" + excel_row_counter.ToString()].Value = tradeResult.TargetPrice;
                ws.Cells["G" + excel_row_counter.ToString()].Value = tradeResult.ExitDate;

                cell = ws.Cells["H" + excel_row_counter.ToString()];
                var exit_link = string.Format(EXIT_FILE, tradeResult.EntryDate.ToString("yyyy_MM_dd"), tradeResult.Ticker);
                cell.Formula = string.Format("=HYPERLINK(\"{0}\", \"{1}\")", exit_link, tradeResult.ExitPrice);
                cell.Style.Font.UnderlineStyle = UnderlineStyle.Single;
                cell.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);

                cell = ws.Cells["I" + excel_row_counter.ToString()];
                var p10_link = string.Format(EXIT_PAST_10, tradeResult.EntryDate.ToString("yyyy_MM_dd"), tradeResult.Ticker);
                cell.Formula = string.Format("=HYPERLINK(\"{0}\", \"{1}\")", p10_link, "P10");
                cell.Style.Font.UnderlineStyle = UnderlineStyle.Single;
                cell.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);

                cell = ws.Cells["J" + excel_row_counter.ToString()];
                var p30_link = string.Format(EXIT_PAST_30, tradeResult.EntryDate.ToString("yyyy_MM_dd"), tradeResult.Ticker);
                cell.Formula = string.Format("=HYPERLINK(\"{0}\", \"{1}\")", p30_link, "P30");
                cell.Style.Font.UnderlineStyle = UnderlineStyle.Single;
                cell.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);

                //if it is winning, hightlight the exist value cell
                //assume the exit price > target price is winning here
                if (tradeResult.ExitPrice > tradeResult.TargetPrice)
                {
                    //this should not override the style set up in the previous section except font color
                    //need to test to confirm
                    ws.Cells["H" + excel_row_counter.ToString()].Style.Font.Color = SpreadsheetColor.FromName(ColorName.Green);
                }

                excel_row_counter++;
            }
            ef.Save(string.Format("tradeResult{0}.xlsx", "Final"));
            Console.WriteLine("Enter any key to exit...");
            Console.ReadKey();

        }
    }
}

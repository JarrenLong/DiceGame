using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace DiceGame
{
  // Game rules: https://public.wsu.edu/~engrmgmt/holt/em530/Docs/DiceGames.htm
  class Program
  {
    static void Main(string[] args)
    {
      int numDays = 10;
      int numWorkers = 6;
      decimal startRandNum = 1;
      decimal endRandNum = 6;
      bool intRandsOnly = true;
      int numIterations = 100;
      Random r = new Random();

      List<Iteration> iterations = new List<Iteration>();
      WorkerSummary ws = null;
      DayWorkerPair dwp = null;
      Iteration round = null;

      // Run specified number of iterations
      for (int it = 0; it < numIterations; it++)
      {
        round = new Iteration();

        for (int d = 0; d < numDays; d++)
        {
          for (int w = 0; w < numWorkers; w++)
          {
            dwp = new DayWorkerPair()
            {
              Map = round.Map,
              DayNumber = d + 1,
              WorkerNumber = w + 1,
              Roll = GetRand(r, startRandNum, endRandNum, intRandsOnly)
            };
            dwp.Calculate();

            round.Map.Add(dwp);
          }
        }

        for (int w = 0; w < numWorkers; w++)
        {
          var wDays = round.Map.Where(x => x.WorkerNumber == w + 1);
          ws = new WorkerSummary()
          {
            WorkerNumber = w + 1,
            TotalRoll = wDays.Sum(x => x.Roll),
            AverageWIP = wDays.Sum(x => x.WIP) / numDays
          };
          ws.AverageRoll = ws.TotalRoll / numDays;

          round.WorkerSummaries.Add(ws);
        }

        // Crunch the numbers for this iteration
        round.TotalRoll = round.Map.Where(x => x.WorkerNumber == numWorkers).Sum(x => x.WIP);
        round.TotalWIP = round.Map.Where(x => x.DayNumber == numDays).Sum(x => x.WIP);
        round.AvgDailyOutput = round.TotalRoll / numDays;
        round.OverallAvgWIP = round.WorkerSummaries.Sum(x => x.AverageWIP) / numWorkers;
        round.OverallAvgRollTotal = round.WorkerSummaries.Sum(x => x.TotalRoll) / numWorkers;
        round.OverallAvgRoll = round.WorkerSummaries.Sum(x => x.AverageRoll) / numWorkers;

        // Random rolls are generated, crunch the numbers for this iteration
        iterations.Add(round);
      }

      // Calculate averages for all iterations run
      round = new Iteration();
      for (int w = 0; w < numWorkers; w++)
        round.WorkerSummaries.Add(new WorkerSummary()
        {
          WorkerNumber = w + 1
        });

      foreach (var it in iterations)
      {
        round.TotalRoll += it.TotalRoll;
        round.TotalWIP += it.TotalWIP;
        round.AvgDailyOutput += it.AvgDailyOutput;
        round.OverallAvgWIP += it.OverallAvgWIP;
        round.OverallAvgRoll += it.OverallAvgRoll;
        round.OverallAvgRollTotal += it.OverallAvgRollTotal;

        for (int w = 0; w < numWorkers; w++)
        {
          ws = round.WorkerSummaries.FirstOrDefault(x => x.WorkerNumber == w + 1);
          var wsb = it.WorkerSummaries.FirstOrDefault(x => x.WorkerNumber == w + 1);

          ws.TotalRoll += wsb.TotalRoll;
          ws.AverageRoll += wsb.AverageRoll;
          ws.AverageWIP += wsb.AverageWIP;
        }
      }

      round.AvgDailyOutput /= numIterations;
      round.OverallAvgWIP /= numIterations;
      round.OverallAvgRoll /= numIterations;
      round.OverallAvgRollTotal /= numIterations;

      for (int w = 0; w < numWorkers; w++)
      {
        ws = round.WorkerSummaries.FirstOrDefault(x => x.WorkerNumber == w + 1);

        ws.TotalRoll /= numIterations;
        ws.AverageRoll /= numIterations;
        ws.AverageWIP /= numIterations;
      }

      // Puke all results out to an Excel spreadsheet and open it
      string outputFile = string.Format("DiceGameOutput_{0:yyyyMMddHHmmss}.xlsx", DateTime.Now);
      Dictionary<string, List<string[]>> excelPages = new Dictionary<string, List<string[]>>();
      List<string[]> sPage = new List<string[]>();
      sPage.Add(new string[] { "Number of Workers:", numWorkers + "" });
      sPage.Add(new string[] { "Number of Days:", numDays + "" });
      sPage.Add(new string[] { "Number of Iterations Run:", numIterations + "" });
      sPage.Add(new string[] { "Dice Type", intRandsOnly ? "Integers only" : "Decimals allowed" });
      sPage.Add(new string[] { "Dice Range:", startRandNum + "-" + endRandNum });
      sPage.Add(new string[] { "" });
      sPage.AddRange(round.ToMapString(numWorkers, numDays));

      excelPages.Add("Master Summary", sPage);

      int i = 0;
      foreach (var it in iterations)
      {
        i++;
        excelPages.Add("Iteration " + i, it.ToMapString(numWorkers, numDays));
      }
      ExportExcel(outputFile, excelPages);
      Process.Start(outputFile);
    }

    private static decimal GetRand(Random r, decimal lo, decimal hi, bool intOnly)
    {
      decimal rnd = lo + ((decimal)r.NextDouble() * (hi - lo));
      if (intOnly)
        rnd = Math.Round(rnd, 0);
      return rnd;
    }

    /// <summary>
    /// Exports a collection of string arrays directly to an Excel spreadsheet
    /// </summary>
    /// <param name="fileName">Path to the file to save</param>
    /// <param name="list">A dictionary of named lists, each key represents a new sheet's title, and the associated value is the data to save in the sheet.</param>
    /// <returns>True on success, else false.</returns>
    private static bool ExportExcel(string fileName, Dictionary<string, List<string[]>> pages)
    {
      if (pages == null || pages.Count == 0)
        return false;

      try
      {
        using (ExcelPackage p = new ExcelPackage())
        {
          // Add all of the defined pages
          foreach (KeyValuePair<string, List<string[]>> pg in pages)
          {
            List<string[]> list = pg.Value;

            // Figure out the max dimensions of the spreadsheet being used
            int rows = list.Count + 1;
            int cols = list.Max(x => x.Length) + 1;

            // Add a worksheet
            ExcelWorksheet ws = p.Workbook.Worksheets.Add(pg.Key);
            // Dump in the data starting in the first cell
            ws.Cells["A1"].LoadFromArrays(list);
            // Autosize the columns
            ws.Cells[ws.Dimension.Address].AutoFitColumns();

            // Roll through and make sure all formulas are put in the correct places, if any are present
            int rc = 1;
            List<int> checkRows = new List<int>();
            foreach (string[] rr in list)
            {
              if (rr.Any(x => x != null && x.StartsWith("=")))
                checkRows.Add(rc);
              rc++;
            }

            string cellVal = "";
            ExcelRange cell = null;
            foreach (int r in checkRows)
            {
              for (int c = 1; c <= cols; c++)
              {
                try
                {
                  cell = ws.Cells[r, c];
                  cellVal = cell.GetValue<string>();
                  if (!string.IsNullOrEmpty(cellVal) && cellVal.StartsWith("="))
                  {
                    cell.Value = "";
                    // NOTE: Trim off leading '=', EPPlus library doesn't like it
                    cell.Formula = cellVal.Substring(1);
                  }
                }
                catch { }
              }
            }
          }

          // Crunch all formulas in the workbook
          p.Workbook.Calculate();
          // Save the workbook
          p.SaveAs(new FileInfo(fileName));
        }
      }
      catch { return false; }

      return true;
    }

  }

  class Iteration
  {
    public List<DayWorkerPair> Map = new List<DayWorkerPair>();
    public List<WorkerSummary> WorkerSummaries = new List<WorkerSummary>();

    public decimal TotalRoll = 0;
    public decimal TotalWIP = 0;
    public decimal AvgDailyOutput = 0;
    public decimal OverallAvgWIP = 0;
    public decimal OverallAvgRoll = 0;
    public decimal OverallAvgRollTotal = 0;

    public List<string[]> ToMapString(int numWorkers, int numDays)
    {
      List<string[]> ret = new List<string[]>();
      List<string> buf = new List<string>();
      List<string> buf2 = new List<string>();

      // Top Row, worker labels
      for (int w = 0; w <= numWorkers; w++)
      {
        if (w == 0)
          buf.Add("");
        else
          buf.AddRange(new string[] { "Worker " + w, "" });
      }
      ret.Add(buf.ToArray());
      buf.Clear();

      // Add in the master table
      int i = 0;
      foreach (var d in Map.GroupBy(x => x.DayNumber))
      {
        buf.Add("Day " + d.Key + " Rolls");
        buf2.Add("Day " + d.Key + " WIPs");

        foreach (var dw in d)
        {
          buf.AddRange(new string[] { dw.Roll.ToString("0.00"), "" });
          buf2.AddRange(new string[] { dw.TempWIP.ToString("0.00"), dw.WIP.ToString("0.00") });
        }

        ret.Add(buf.ToArray());
        ret.Add(buf2.ToArray());
        buf.Clear();
        buf2.Clear();

        i++;
      }

      ret.Add(new string[] { "" });

      // Add the worker summaries
      buf.Add("");
      var tmp = WorkerSummaries.Select(x => new string[] { "Total Roll:", x.TotalRoll.ToString("0.00") });
      foreach (var t in tmp)
        buf.AddRange(t);
      ret.Add(buf.ToArray());
      buf.Clear();

      buf.Add("");
      tmp = WorkerSummaries.Select(x => new string[] { "Avg Roll:", x.AverageRoll.ToString("0.00") });
      foreach (var t in tmp)
        buf.AddRange(t);
      ret.Add(buf.ToArray());
      buf.Clear();

      buf.Add("");
      tmp = WorkerSummaries.Select(x => new string[] { "Avg WIP:", x.AverageWIP.ToString("0.00") });
      foreach (var t in tmp)
        buf.AddRange(t);
      ret.Add(buf.ToArray());
      buf.Clear();

      ret.Add(new string[] { "" });

      //if (buf.Count > 0)
      //{
      //  ret.Add(buf.ToArray());

      //  ret.Add(new string[] { "" });
      //}
      ret.Add(new string[] { "Total Roll:", TotalRoll.ToString("0.00") });
      ret.Add(new string[] { "Total WIP:", TotalWIP.ToString("0.00") });
      ret.Add(new string[] { "Avg. Daily Output:", AvgDailyOutput.ToString("0.00") });
      ret.Add(new string[] { "Avg. WIP:", OverallAvgWIP.ToString("0.00") });
      ret.Add(new string[] { "Avg. Roll:", OverallAvgRoll.ToString("0.00") });
      ret.Add(new string[] { "Avg. Roll. Total:", OverallAvgRollTotal.ToString("0.00") });

      return ret;
    }
  }

  class DayWorkerPair
  {
    public List<DayWorkerPair> Map;

    public int DayNumber = 0;
    public int WorkerNumber = 0;
    public decimal Roll = 0;
    public decimal TempWIP = -1;
    public decimal WIP = -1;

    private DayWorkerPair GetCalced(Func<DayWorkerPair, bool> q)
    {
      DayWorkerPair buf = Map.FirstOrDefault(x => x.DayNumber == DayNumber && x.WorkerNumber == WorkerNumber - 1);
      if (buf != null && (buf.TempWIP == -1 || buf.WIP == -1))
        buf.Calculate();

      return buf;
    }

    public void Calculate()
    {
      decimal lastWorkerCap = GetCalced(x => x.DayNumber == DayNumber && x.WorkerNumber == WorkerNumber - 1)?.TempWIP ?? Roll;
      decimal thisWorkerYesterdayWIP = GetCalced(x => x.DayNumber == DayNumber - 1 && x.WorkerNumber == WorkerNumber)?.WIP ?? 0;
      decimal nextWorkerCap = GetCalced(x => x.DayNumber == DayNumber && x.WorkerNumber == WorkerNumber + 1)?.Roll ?? 0;

      TempWIP = Math.Min(lastWorkerCap, Roll) + thisWorkerYesterdayWIP;
      WIP = nextWorkerCap >= TempWIP ? 0 : TempWIP - nextWorkerCap;
    }
  }

  class WorkerSummary
  {
    public int WorkerNumber;
    public decimal TotalRoll;
    public decimal AverageRoll;
    public decimal AverageWIP;
  }
}

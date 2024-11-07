using System.Text.Json;
using Dapper;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace ConsoleApp1;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var mssqlConnectionString = "Data Source=dbz.siemens.com;Trust Server Certificate=True;Authentication=ActiveDirectoryInteractive";
        using var connection = new SqlConnection(mssqlConnectionString);
        
        string sql = File.ReadAllText("sql1.sql");
        string sql2 = File.ReadAllText("sql2.sql");
        string sql3 = File.ReadAllText("sql3.sql");


     
        connection.Execute("USE importdb");
        //var data = connection.Query<Sql1>(sql).ToList();
        var data = connection.Query<Sql2>(sql2).ToList();
        var data2 = connection.Query<Sql3>(sql3).ToList();

        //var data = JsonConvert.DeserializeObject<List<Sql1>>(dataJson);
        
         using (var package = new ExcelPackage())
        {
            // První list - původní data a sloupcový graf
            var worksheet = package.Workbook.Worksheets.Add("Data");

            // Přidání hlavičky
            worksheet.Cells["A1"].Value = "Uživatel";
            worksheet.Cells["B1"].Value = "Taby";

            // Přidání dat
            var groupedData = data.GroupBy(d => d.NTUserName)
                .Select(g => new { Uzivatel = g.Key, Taby = g.Sum(d => d.Count) })
                .OrderByDescending(g => g.Taby)  // Seřazení dat sestupně podle počtu tabů
                .ToList();

            for (int i = 0; i < groupedData.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = groupedData[i].Uzivatel;
                worksheet.Cells[i + 2, 2].Value = groupedData[i].Taby;
            }

            // Vytvoření sloupcového grafu
            var columnChart = worksheet.Drawings.AddChart("Graf", eChartType.ColumnClustered);
            columnChart.Title.Text = "Počet tabů dohromady otevřený daným uživatelem";
            columnChart.SetPosition(1, 0, 5, 0);
            columnChart.SetSize(800, 400);

            // Nastavení dat pro sloupcový graf
            var columnSeries = columnChart.Series.Add(worksheet.Cells[2, 2, groupedData.Count + 1, 2], worksheet.Cells[2, 1, groupedData.Count + 1, 1]);
            columnSeries.Header = "Počet tabů";

            // Druhý list - koláčové grafy pro každého uživatele
            var userWorksheet = package.Workbook.Worksheets.Add("Uživatelé");

            int currentRow = 1;

            foreach (var userGroup in data.GroupBy(d => d.NTUserName))
            {
                string userName = userGroup.Key;

                // Přidání dat pro aktuálního uživatele
                userWorksheet.Cells[currentRow, 1].Value = "Uživatel";
                userWorksheet.Cells[currentRow, 2].Value = userName;
                currentRow++;

                userWorksheet.Cells[currentRow, 1].Value = "Název Tabu";
                userWorksheet.Cells[currentRow, 2].Value = "Počet";
                currentRow++;

                int startRow = currentRow;
                foreach (var item in userGroup)
                {
                    userWorksheet.Cells[currentRow, 1].Value = item.MappedValue;
                    userWorksheet.Cells[currentRow, 2].Value = item.Count;
                    currentRow++;
                }
                int endRow = currentRow - 1;

                // Vytvoření koláčového grafu pro aktuálního uživatele
                var pieChart = userWorksheet.Drawings.AddChart($"Graf_{userName}", eChartType.Pie);
                pieChart.Title.Text = $"Taby použité uživatelem {userName}";
                pieChart.SetPosition(startRow - 2, 0, 3, 0);
                pieChart.SetSize(500, 300);

// Nastavení dat pro koláčový graf
                var pieSeries = pieChart.Series.Add(userWorksheet.Cells[startRow, 2, endRow, 2], userWorksheet.Cells[startRow, 1, endRow, 1]);
                pieSeries.Header = "Počet použití";

// Nastavení zobrazení hodnot
               // pieChart.XAxis.DisplayUnit = 1;
               // pieChart.XAxis.Format = "#,##0";

// Nastavení legendy pro zobrazení MappedValue
                pieChart.Legend.Position = eLegendPosition.Right;

                
                
// Přidání popisků dat (procenta)
              //  pieChart.PlotArea.DataTable.ShowKeys = true;

                currentRow += 20; // Větší mezera mezi grafy pro lepší čitelnost
            }
            
            // Třetí list - Histogram
            var histogramWorksheet = package.Workbook.Worksheets.Add("Frekvence využití");

            // Přidání hlavičky
            histogramWorksheet.Cells["A1"].Value = "Hodina";
            histogramWorksheet.Cells["B1"].Value = "Počet Procedur";

            // Seskupení dat podle hodin
            var groupedData2 = data2.GroupBy(d => 
                {
                    // Předpokládáme, že TimeOnly je ve formátu "HH:mm:ss"
                    var timeParts = d.TimeOnly.Split(':');
                    return int.Parse(timeParts[0]); // Vrátí hodinu jako int
                })
                .Select(g => new { Hodina = g.Key, PocetProcedur = g.Count() })
                .OrderBy(g => g.Hodina)
                .ToList();

            for (int i = 0; i < groupedData2.Count; i++)
            {
                histogramWorksheet.Cells[i + 2, 1].Value = groupedData2[i].Hodina;
                histogramWorksheet.Cells[i + 2, 2].Value = groupedData2[i].PocetProcedur;
            }

            // Vytvoření histogramu
            var histogramChart = histogramWorksheet.Drawings.AddChart("ZatíženostServeruGraf", eChartType.ColumnClustered);
            histogramChart.Title.Text = "Zatíženost Serveru";
            histogramChart.SetPosition(1, 0, 3, 0);
            histogramChart.SetSize(800, 400);

            // Nastavení dat pro histogram
            var histogramSeries = histogramChart.Series.Add(histogramWorksheet.Cells[2, 2, groupedData2.Count + 1, 2], histogramWorksheet.Cells[2, 1, groupedData2.Count + 1, 1]);
            histogramSeries.Header = "Počet Procedur";

         
            
            
            
            // Čtvrtý list - Rozpady použití karet
var breakdownWorksheet2 = package.Workbook.Worksheets.Add("Počty použití tabů");

// Přidání hlavičky
breakdownWorksheet2.Cells["A1"].Value = "Kategorie";
breakdownWorksheet2.Cells["B1"].Value = "Počet";

// Zpracování dat pro kategorie
var categories2 = data2.GroupBy(d => d.MappedValue.Split('-')[0].Trim())
    .Select(g => new { Kategorie = g.Key, Pocet = g.Count() })
    .OrderByDescending(g => g.Pocet)
    .ToList();

int currentRow2 = 2;
for (int i = 0; i < categories2.Count; i++)
{
    breakdownWorksheet2.Cells[currentRow2, 1].Value = categories2[i].Kategorie;
    breakdownWorksheet2.Cells[currentRow2, 2].Value = categories2[i].Pocet;
    currentRow2++;
}

// Vytvoření koláčového grafu pro kategorie
var categoryPieChart = breakdownWorksheet2.Drawings.AddChart("KategorieGraf", eChartType.Pie);
categoryPieChart.Title.Text = "Rozpady použití karet podle kategorií";
categoryPieChart.SetPosition(1, 0, 3, 0);
categoryPieChart.SetSize(500, 300);

var categorySeries = categoryPieChart.Series.Add(breakdownWorksheet2.Cells[2, 2, categories2.Count + 1, 2], breakdownWorksheet2.Cells[2, 1, categories2.Count + 1, 1]);
categorySeries.Header = "Počet použití";

// Přidání mezery mezi grafy
currentRow2 += 20;

// Zpracování dat pro standardní vs Info plus
var standardVsInfoPlus = data2.GroupBy(d => d.MappedValue.Contains("Info plus") ? "Info plus" : "Standardní")
    .Select(g => new { Typ = g.Key, Pocet = g.Count() })
    .OrderByDescending(g => g.Pocet)
    .ToList();

int startRow2 = currentRow2;
breakdownWorksheet2.Cells[startRow2, 1].Value = "Typ";
breakdownWorksheet2.Cells[startRow2, 2].Value = "Počet";

for (int i = 0; i < standardVsInfoPlus.Count; i++)
{
    breakdownWorksheet2.Cells[startRow2 + i + 1, 1].Value = standardVsInfoPlus[i].Typ;
    breakdownWorksheet2.Cells[startRow2 + i + 1, 2].Value = standardVsInfoPlus[i].Pocet;
}

// Vytvoření koláčového grafu pro standardní vs Info plus
var standardVsInfoPlusPieChart = breakdownWorksheet2.Drawings.AddChart("StandardVsInfoPlusGraf", eChartType.Pie);
standardVsInfoPlusPieChart.Title.Text = "Standardní vs Info plus";
standardVsInfoPlusPieChart.SetPosition(startRow2 + standardVsInfoPlus.Count + 2, 0, 3, 0);
standardVsInfoPlusPieChart.SetSize(500, 300);

var standardVsInfoPlusSeries = standardVsInfoPlusPieChart.Series.Add(breakdownWorksheet2.Cells[startRow2 + 1, 2, startRow2 + standardVsInfoPlus.Count, 2], breakdownWorksheet2.Cells[startRow2 + 1, 1, startRow2 + standardVsInfoPlus.Count, 1]);
standardVsInfoPlusSeries.Header = "Počet použití";

            //////////////////////////////////////////
            /// pokus



            // Přidání mezery mezi grafy
            currentRow2 += 20;

            // Zpracování dat pro kategorie v plném detailu
            var detailedCategories = data2.GroupBy(d => d.MappedValue.Trim())
                .Select(g => new { Kategorie = g.Key, Pocet = g.Count() })
                .OrderByDescending(g => g.Pocet)
                .ToList();

            // Přidání dat pro detailní kategorie do worksheetu
            int detailedStartRow = currentRow2;
            breakdownWorksheet2.Cells[detailedStartRow, 1].Value = "Detailní Kategorie";
            breakdownWorksheet2.Cells[detailedStartRow, 2].Value = "Počet";

            for (int i = 0; i < detailedCategories.Count; i++)
            {
                breakdownWorksheet2.Cells[detailedStartRow + i + 1, 1].Value = detailedCategories[i].Kategorie;
                breakdownWorksheet2.Cells[detailedStartRow + i + 1, 2].Value = detailedCategories[i].Pocet;
            }

            // Vytvoření koláčového grafu pro detailní kategorie
            var detailedCategoryPieChart = breakdownWorksheet2.Drawings.AddChart("DetailniKategorieGraf", eChartType.Pie);
            detailedCategoryPieChart.Title.Text = "Rozpady použití karet podle detailních kategorií";
            detailedCategoryPieChart.SetPosition(detailedStartRow + detailedCategories.Count + 2, 0, 3, 0);
            detailedCategoryPieChart.SetSize(500, 300);

            var detailedCategorySeries = detailedCategoryPieChart.Series.Add(breakdownWorksheet2.Cells[detailedStartRow + 1, 2, detailedStartRow + detailedCategories.Count, 2], breakdownWorksheet2.Cells[detailedStartRow + 1, 1, detailedStartRow + detailedCategories.Count, 1]);
            detailedCategorySeries.Header = "Počet použití";




            ///////////////////










            // paty list - Rozpady použití karet
            var breakdownWorksheet = package.Workbook.Worksheets.Add("Rozpady kategorií tabů");

            // Zpracování dat pro každou kategorii
            var categories = data2.GroupBy(d => d.MappedValue.Split('-')[0].Trim())
                .Select(g => new { Kategorie = g.Key, Data = g.ToList() })
                .ToList();

            int currentRow3 = 1;

            foreach (var category in categories)
            {
                // Přidání hlavičky pro aktuální kategorii
                breakdownWorksheet.Cells[currentRow3, 1].Value = "Kategorie";
                breakdownWorksheet.Cells[currentRow3, 2].Value = "Počet";
                currentRow3++;

                // Zpracování dat pro aktuální kategorii
                var categoryData = category.Data.GroupBy(d => d.MappedValue)
                    .Select(g => new { MappedValue = g.Key, Pocet = g.Count() })
                    .OrderByDescending(g => g.Pocet)
                    .ToList();

                int startRow = currentRow3;
                for (int i = 0; i < categoryData.Count; i++)
                {
                    breakdownWorksheet.Cells[currentRow3, 1].Value = categoryData[i].MappedValue;
                    breakdownWorksheet.Cells[currentRow3, 2].Value = categoryData[i].Pocet;
                    currentRow3++;
                }
                int endRow = currentRow3 - 1;

                // Vytvoření koláčového grafu pro aktuální kategorii
                var pieChart = breakdownWorksheet.Drawings.AddChart($"Graf_{category.Kategorie}", eChartType.Pie);
                pieChart.Title.Text = $"Rozpad pro kategorii {category.Kategorie}";
                pieChart.SetPosition(startRow - 2, 0, 3, 0);
                pieChart.SetSize(500, 300);

                var pieSeries = pieChart.Series.Add(breakdownWorksheet.Cells[startRow, 2, endRow, 2], breakdownWorksheet.Cells[startRow, 1, endRow, 1]);
                pieSeries.Header = "Počet použití";

                currentRow3 += 20; // Větší mezera mezi grafy pro lepší čitelnost
            }
            
            
            // Uložení souboru
            File.WriteAllBytes("vystup.xlsx", package.GetAsByteArray());
        }

        
        int z = 0;
        Console.WriteLine("Hello, World!");
    }
}
using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CSharpFlexGrid
{
    public class BookingOrderReader
    {
        private DataSet _excelData;

        public string ProcessBookingOrder(string filePath)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        _excelData = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = false }
                        });
                    }
                }

                // Find the correct sheet with booking data
                var sheet = FindBookingOrderSheet(_excelData);
                if (sheet == null)
                {
                    throw new Exception("No booking order sheet found. Looking for sheet with substantial data (>10 rows, >10 columns).");
                }

                var bookingOrder = new BookingOrder();

                // Check for a unique marker to determine the file format
                if (FindRowIndex(sheet, "Programs Name") != -1)
                {
                    ParseDirectNCCAL(sheet, bookingOrder);
                }
                else
                {
                    ParseAlMulla(sheet, bookingOrder);
                }

                var allDates = bookingOrder.Spots.SelectMany(s => s.Dates)
                                                  .Select(d => DateTime.Parse(d))
                                                  .OrderBy(d => d)
                                                  .ToList();
                if (allDates.Any())
                {
                    bookingOrder.CampaignPeriod.StartDate = allDates.First().ToString("yyyy-MM-dd");
                    bookingOrder.CampaignPeriod.EndDate = allDates.Last().ToString("yyyy-MM-dd");
                }

                if (bookingOrder.TotalSpots == 0)
                {
                    bookingOrder.TotalSpots = bookingOrder.Spots.Sum(s => s.TotalSpots);
                }

                return JsonConvert.SerializeObject(bookingOrder, Formatting.Indented);
            }
            catch (Exception ex)
            {
                return JsonConvert.SerializeObject(new { error = ex.Message, stackTrace = ex.StackTrace });
            }
        }

        private DataTable FindBookingOrderSheet(DataSet dataSet)
        {
            // Strategy: Find the sheet with the most data (rows and columns)
            // Booking orders typically have 20+ rows and 10+ columns
            // Skip sheets with minimal data (like reference tables)

            DataTable bestSheet = null;
            int bestScore = 0;

            foreach (DataTable table in dataSet.Tables)
            {
                // Score based on rows and columns
                // Require minimum 10 rows and 10 columns to be considered
                if (table.Rows.Count >= 10 && table.Columns.Count >= 10)
                {
                    int score = table.Rows.Count * table.Columns.Count;
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestSheet = table;
                    }
                }
            }

            // If no sheet meets criteria, fall back to first sheet
            return bestSheet ?? (dataSet.Tables.Count > 0 ? dataSet.Tables[0] : null);
        }

        private void ParseAlMulla(DataTable sheet, BookingOrder bookingOrder)
        {
            bookingOrder.CompanyName = FindValueInRow(sheet, "Company Name:", "Agency:", "Client:", "Client Name");
            bookingOrder.Advertiser = FindValueInRow(sheet, "ADVERTISER:", "Client");
            bookingOrder.Product = FindValueInRow(sheet, "PRODUCT:", "Campaign");
            bookingOrder.Agency = bookingOrder.CompanyName;

            var costStr = FindValueBelow(sheet, "Package Cost", "Total Cost", "Net Cost", "Gross Cost");
            if (decimal.TryParse(costStr, out var grossCost))
            {
                bookingOrder.GrossCost = grossCost;
            }

            var spotsStr = FindValueBelow(sheet, "Number of Spots:", "Total Spots");
            if (int.TryParse(spotsStr, out var totalSpots))
            {
                bookingOrder.TotalSpots = totalSpots;
            }

            ParseAlMullaSchedule(sheet, bookingOrder);
        }

        private void ParseDirectNCCAL(DataTable sheet, BookingOrder bookingOrder)
        {
            bookingOrder.Agency = FindValueInRow(sheet, "Agency :");
            bookingOrder.Advertiser = FindValueInRow(sheet, "Advertiser :");
            bookingOrder.Product = FindValueInRow(sheet, "Product :");
            bookingOrder.CompanyName = bookingOrder.Agency;

            var costStr = FindValueInRow(sheet, "Package Cost:");
            if (decimal.TryParse(costStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var grossCost))
            {
                bookingOrder.GrossCost = grossCost;
            }

            var spotsStr = FindValueInRow(sheet, "Total Spots");
            if (int.TryParse(spotsStr, out var totalSpots))
            {
                bookingOrder.TotalSpots = totalSpots;
            }

            ParseDirectNCCALSchedule(sheet, bookingOrder);
        }

        private void ParseAlMullaSchedule(DataTable sheet, BookingOrder bookingOrder)
        {
            var (year, month) = FindYearAndMonth(sheet);
            var headerRowIndex = FindRowIndex(sheet, "PROGRAMS", "PROGRAMMES", "PROGRAMME", "Schedule", "Program Name");
            if (headerRowIndex == -1) throw new Exception("AlMulla: Header row with 'PROGRAMS' or similar not found.");

            var calendarRowIndex = headerRowIndex - 1;
            if (calendarRowIndex < 0) throw new Exception("AlMulla: Calendar row not found (expected above header row).");

            int programCol = FindColumnIndex(sheet, headerRowIndex, "PROGRAMS", "PROGRAMMES", "PROGRAMME", "Program Name", "Show");
            if (programCol == -1) throw new Exception("AlMulla: Column 'PROGRAMS' or similar not found.");

            int timeCol = FindColumnIndex(sheet, headerRowIndex, "TIME", "START TIME", "Start", "Slot");
            if (timeCol == -1) throw new Exception("AlMulla: Column 'TIME' or similar not found.");

            int durCol = FindColumnIndex(sheet, headerRowIndex, "DUR", "DURATION", "Length", "Secs");
            if (durCol == -1) throw new Exception("AlMulla: Column 'DUR' or similar not found.");
            
            int firstDayCol = -1;
            for(int j=0; j<sheet.Columns.Count; j++)
            {
                var cellObj = sheet.Rows[calendarRowIndex][j];
                var cellValue = cellObj?.ToString();

                if(!string.IsNullOrWhiteSpace(cellValue))
                {
                    // Check if it's a number
                    if(int.TryParse(cellValue, out _))
                    {
                        firstDayCol = j;
                        System.Diagnostics.Debug.WriteLine($"DEBUG: First day column found at index {j}, calendar row {calendarRowIndex}, value: {cellValue}");
                        break;
                    }
                    // Check if it's a DateTime (day numbers might be read as dates)
                    else if(cellObj is DateTime || DateTime.TryParse(cellValue, out _))
                    {
                        firstDayCol = j;
                        System.Diagnostics.Debug.WriteLine($"DEBUG: First day column found at index {j}, calendar row {calendarRowIndex}, value: {cellValue} (DateTime)");
                        break;
                    }
                }
            }
            
            if(firstDayCol == -1) throw new Exception("AlMulla: First day column not found in calendar row.");

            int totalCol = FindColumnIndex(sheet, calendarRowIndex, "Total", "TOTAL SPOTS");
            if (totalCol == -1) totalCol = sheet.Columns.Count;


            // Pre-build a list of resolved dates for each day column to handle multi-month campaigns
            int currentMonth = month;
            int currentYear = year;
            int prevDay = 0;
            var columnDates = new DateTime?[sheet.Columns.Count];

            for (int c = firstDayCol; c < totalCol; c++)
            {
                var cellObj = sheet.Rows[calendarRowIndex][c];
                var dayHeader = cellObj?.ToString();
                if (string.IsNullOrWhiteSpace(dayHeader)) continue;

                int day = 0;

                // If Excel stored it as a DateTime, use its month and day directly
                if (cellObj is DateTime dtVal)
                {
                    day = dtVal.Day;
                    // Trust the month from the DateTime if it looks valid (not the 1899/1900 epoch)
                    if (dtVal.Year > 1900)
                    {
                        columnDates[c] = dtVal;
                        prevDay = day;
                        continue;
                    }
                }
                else if (int.TryParse(dayHeader, out day))
                {
                    // Day is an integer
                }
                else if (DateTime.TryParse(dayHeader, out var dateValue))
                {
                    day = dateValue.Day;
                    if (dateValue.Year > 1900)
                    {
                        columnDates[c] = dateValue;
                        prevDay = day;
                        continue;
                    }
                }

                if (day > 0)
                {
                    // Detect month rollover: day number decreased significantly
                    if (prevDay > 0 && day < prevDay && prevDay > 15)
                    {
                        currentMonth++;
                        if (currentMonth > 12)
                        {
                            currentMonth = 1;
                            currentYear++;
                        }
                    }
                    prevDay = day;

                    try
                    {
                        columnDates[c] = new DateTime(currentYear, currentMonth, day);
                    }
                    catch { /* Ignore invalid dates */ }
                }
            }

            for (int r = headerRowIndex + 1; r < sheet.Rows.Count; r++)
            {
                var programName = sheet.Rows[r][programCol]?.ToString();
                if (string.IsNullOrWhiteSpace(programName) || programName.ToUpper().Contains("PAGE TOTAL") || programName.ToUpper().Contains("TOTAL"))
                {
                    if (programName?.ToUpper().Contains("TOTAL") == true) break;
                    continue;
                }

                var timeStr = sheet.Rows[r][timeCol]?.ToString();
                if (DateTime.TryParse(timeStr, out var timeVal))
                {
                    timeStr = timeVal.ToString("HH:mm");
                }

                var spot = new Spot
                {
                    ProgrammeName = programName,
                    ProgrammeStartTime = timeStr,
                    Duration = sheet.Rows[r][durCol]?.ToString()
                };

                for (int c = firstDayCol; c < totalCol; c++)
                {
                    var spotMarker = sheet.Rows[r][c]?.ToString();

                    if (!string.IsNullOrWhiteSpace(spotMarker) && columnDates[c].HasValue)
                    {
                        var date = columnDates[c].Value;

                        // Parse spot marker to get count
                        int spotCount = 0;
                        if (int.TryParse(spotMarker, out var parsedCount))
                        {
                            spotCount = parsedCount;
                        }

                        // Add date multiple times based on spot count (skip if 0)
                        if (spotCount > 0)
                        {
                            for (int i = 0; i < spotCount; i++)
                            {
                                spot.Dates.Add(date.ToString("yyyy-MM-dd"));
                            }
                        }
                    }
                }

                if (totalCol < sheet.Columns.Count && int.TryParse(sheet.Rows[r][totalCol]?.ToString(), out int spotTotal))
                {
                    spot.TotalSpots = spotTotal;
                } else {
                    spot.TotalSpots = spot.Dates.Count;
                }

                if (spot.Dates.Any())
                {
                    bookingOrder.Spots.Add(spot);
                }
            }
        }

        private void ParseDirectNCCALSchedule(DataTable sheet, BookingOrder bookingOrder)
        {
            var (year, month) = FindYearAndMonth(sheet);
            var headerRowIndex = FindRowIndex(sheet, "Programs Name");
            if (headerRowIndex == -1) throw new Exception("DirectNCCAL: Header row with 'Programs Name' not found.");

            var calendarRowIndex = headerRowIndex;

            int programCol = FindColumnIndex(sheet, headerRowIndex, "Programs Name");
            if (programCol == -1) throw new Exception("DirectNCCAL: Column 'Programs Name' not found.");

            int timeCol = FindColumnIndex(sheet, headerRowIndex, "Time (KWT)");
            if (timeCol == -1) throw new Exception("DirectNCCAL: Column 'Time (KWT)' not found.");

            int priceCol = FindColumnIndex(sheet, headerRowIndex, "Price in US $");
            if (priceCol == -1) throw new Exception("DirectNCCAL: Column 'Price in US $' not found.");

            int totalCol = FindColumnIndex(sheet, headerRowIndex, "Total Spots");
            if (totalCol == -1) throw new Exception("DirectNCCAL: Column 'Total Spots' not found.");
            
            int firstDayCol = priceCol + 1;

            // Pre-build resolved dates for each column to handle multi-month and DateTime cells
            int currentMonth = month;
            int currentYear = year;
            int prevDay = 0;
            var columnDates = new DateTime?[sheet.Columns.Count];

            for (int c = firstDayCol; c < totalCol; c++)
            {
                var cellObj = sheet.Rows[calendarRowIndex][c];
                var dayHeader = cellObj?.ToString();
                if (string.IsNullOrWhiteSpace(dayHeader)) continue;

                // If Excel stored it as a DateTime object, use it directly
                if (cellObj is DateTime dtVal)
                {
                    columnDates[c] = dtVal;
                    prevDay = dtVal.Day;
                    continue;
                }

                // Try parsing as full DateTime string
                if (DateTime.TryParse(dayHeader, out var parsedDate))
                {
                    columnDates[c] = parsedDate;
                    prevDay = parsedDate.Day;
                    continue;
                }

                // Fall back to parsing as day number with month rollover detection
                if (int.TryParse(dayHeader, out int day) && day > 0)
                {
                    if (prevDay > 0 && day < prevDay && prevDay > 15)
                    {
                        currentMonth++;
                        if (currentMonth > 12)
                        {
                            currentMonth = 1;
                            currentYear++;
                        }
                    }
                    prevDay = day;

                    try
                    {
                        columnDates[c] = new DateTime(currentYear, currentMonth, day);
                    }
                    catch { }
                }
            }

            for (int r = headerRowIndex + 1; r < sheet.Rows.Count; r++)
            {
                var programName = sheet.Rows[r][programCol]?.ToString();
                if (string.IsNullOrWhiteSpace(programName) || programName.ToUpper().Contains("TOTAL"))
                {
                    if (programName?.ToUpper().Contains("TOTAL") == true) break;
                    continue;
                }

                var timeStr = sheet.Rows[r][timeCol]?.ToString();
                if (DateTime.TryParse(timeStr, out var timeVal))
                {
                    timeStr = timeVal.ToString("HH:mm");
                }

                var spot = new Spot
                {
                    ProgrammeName = programName,
                    ProgrammeStartTime = timeStr,
                    Duration = "41"
                };

                for (int c = firstDayCol; c < totalCol; c++)
                {
                    var spotMarker = sheet.Rows[r][c]?.ToString();

                    if (!string.IsNullOrWhiteSpace(spotMarker) && columnDates[c].HasValue)
                    {
                        if (int.TryParse(spotMarker, out int spotCount))
                        {
                            for (int i = 0; i < spotCount; i++)
                            {
                                spot.Dates.Add(columnDates[c].Value.ToString("yyyy-MM-dd"));
                            }
                        }
                    }
                }

                if (totalCol < sheet.Columns.Count && int.TryParse(sheet.Rows[r][totalCol]?.ToString(), out int spotTotal))
                {
                    spot.TotalSpots = spotTotal;
                }
                else
                {
                    spot.TotalSpots = spot.Dates.Count;
                }

                if (spot.Dates.Any())
                {
                    bookingOrder.Spots.Add(spot);
                }
            }
        }

        private string FindValueInRow(DataTable sheet, params string[] labels)
        {
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                for (int j = 0; j < sheet.Columns.Count; j++)
                {
                    var cellValue = sheet.Rows[i][j]?.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(cellValue) && labels.Any(label => cellValue.IndexOf(label, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        for (int k = j + 1; k < sheet.Columns.Count; k++)
                        {
                            var value = sheet.Rows[i][k]?.ToString();
                            if (!string.IsNullOrWhiteSpace(value)) return value.Trim();
                        }
                    }
                }
            }
            return null;
        }
        
        private string FindValueBelow(DataTable sheet, params string[] labels)
        {
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                for (int j = 0; j < sheet.Columns.Count; j++)
                {
                    var cellValue = sheet.Rows[i][j]?.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(cellValue) && labels.Any(label => cellValue.IndexOf(label, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        for (int r = i + 1; r < i + 5 && r < sheet.Rows.Count; r++)
                        {
                            var value = sheet.Rows[r][j]?.ToString();
                            if (!string.IsNullOrWhiteSpace(value)) return value.Trim();
                        }
                    }
                }
            }
            return null;
        }

        private (int year, int month) FindYearAndMonth(DataTable sheet)
        {
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            bool yearFound = false;
            bool monthFound = false;

            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                for (int j = 0; j < sheet.Columns.Count; j++)
                {
                    var cellValue = sheet.Rows[i][j]?.ToString();
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        if (!yearFound)
                        {
                            var match = Regex.Match(cellValue, @"\b(20\d{2})\b");
                            if (match.Success)
                            {
                                year = int.Parse(match.Value);
                                yearFound = true;
                            }
                        }

                        if (!monthFound)
                        {
                            for (int m = 1; m <= 12; m++)
                            {
                                if (cellValue.ToUpper().Contains(CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(m).ToUpper()) ||
                                    cellValue.ToUpper().Contains(CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName(m).ToUpper()))
                                {
                                    month = m;
                                    monthFound = true;
                                    break;
                                }
                            }
                        }

                        if (yearFound && monthFound) return (year, month);
                    }
                }
            }
            return (year, month);
        }

        private int FindRowIndex(DataTable sheet, params string[] contents)
        {
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                var rowText = string.Join(" ", sheet.Rows[i].ItemArray.Select(c => c?.ToString() ?? ""));
                if (contents.Any(content => rowText.IndexOf(content, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return i;
                }
            }
            return -1;
        }

        private int FindColumnIndex(DataTable sheet, int headerRow, params string[] columnNames)
        {
            if (headerRow < 0 || headerRow >= sheet.Rows.Count) return -1;
            for (int j = 0; j < sheet.Columns.Count; j++)
            {
                var cellValue = Regex.Replace(sheet.Rows[headerRow][j]?.ToString() ?? "", @"\s+", " ").Trim();
                if (!string.IsNullOrWhiteSpace(cellValue) && columnNames.Any(columnName => cellValue.IndexOf(columnName, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return j;
                }
            }
            return -1;
        }
    }
}
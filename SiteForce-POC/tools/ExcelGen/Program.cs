using ClosedXML.Excel;

var wb = new XLWorkbook();
var ws = wb.Worksheets.Add("Attendance");

ws.Cell(1, 1).Value = "WorkerId";
ws.Cell(1, 2).Value = "Site";
ws.Cell(1, 3).Value = "DaysPresent";
ws.Cell(1, 4).Value = "DayRate";

var sites = new[]
{
    "Mumbai-Andheri", "Bangalore-Whitefield", "Chennai-OMR",
    "Hyderabad-Gachibowli", "Delhi-Gurugram", "Pune-Hinjewadi"
};

var random = new Random(42);

var siteRateRanges = new Dictionary<string, (int min, int max)>
{
    ["Mumbai-Andheri"] = (750, 1200),
    ["Bangalore-Whitefield"] = (700, 1100),
    ["Chennai-OMR"] = (650, 1000),
    ["Hyderabad-Gachibowli"] = (600, 950),
    ["Delhi-Gurugram"] = (950, 1150),
    ["Pune-Hinjewadi"] = (900, 1050),
};

int row = 2;
int workerSeq = 10000001;

// Sites where all workers should have high attendance (no disputes)
var noDisputeSites = new HashSet<string> { "Delhi-Gurugram", "Pune-Hinjewadi" };

foreach (var site in sites)
{
    var (minRate, maxRate) = siteRateRanges[site];

    for (int w = 0; w < 200; w++)
    {
        var workerId = workerSeq.ToString();
        workerSeq++;

        int daysPresent;
        if (noDisputeSites.Contains(site))
        {
            // Ensure high attendance so net pay stays above dispute threshold
            daysPresent = random.Next(22, 29);
        }
        else if (random.NextDouble() < 0.12)
        {
            daysPresent = random.Next(3, 10);
        }
        else if (random.NextDouble() < 0.2)
        {
            daysPresent = random.Next(10, 18);
        }
        else
        {
            daysPresent = random.Next(18, 29);
        }

        var dayRate = random.Next(minRate, maxRate + 1);

        ws.Cell(row, 1).Value = workerId;
        ws.Cell(row, 2).Value = site;
        ws.Cell(row, 3).Value = daysPresent;
        ws.Cell(row, 4).Value = dayRate;
        row++;
    }
}

var path = Path.Combine("..", "SampleAttendance.xlsx");
wb.SaveAs(path);
Console.WriteLine($"Created: {Path.GetFullPath(path)} with {row - 2} workers across {sites.Length} sites");

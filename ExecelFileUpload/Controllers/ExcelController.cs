using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("[controller]")]
public class ExcelController : ControllerBase
{
    private ILogger<ExcelController> _ilogger;

    public ExcelController(ILogger<ExcelController> logger)
    {
        _ilogger = logger;
    }

    [HttpPost("ReadExcelFrom_FormData")]
    public async Task<IActionResult> ReadExcelFrom_FormData(IFormFile file, CancellationToken cancellationToken)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("File is empty.");
        }

        var data = new List<Dictionary<string, string>>();

        using (var stream = new MemoryStream())
        {
            await file.CopyToAsync(stream, cancellationToken);
            stream.Position = 0; // Reset the stream position to the beginning

            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet.RangeUsed().RowsUsed();

                var headerRow = rows.First(); // Assumes the first row is the header row
                var headers = headerRow.Cells().Select(c => c.Value.ToString()).ToList();

                foreach (var row in rows.Skip(1))
                {
                    var rowData = new Dictionary<string, string>();
                    foreach (var cell in row.Cells())
                    {
                        var header = headers[cell.Address.ColumnNumber - 1];
                        rowData[header] = cell.Value.ToString();
                    }
                    data.Add(rowData);
                }
            }
        }

        return Ok(data);
    }
}
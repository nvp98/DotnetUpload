using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace UploadExcelAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UploadFileController : ControllerBase
    {
        private ILogger<UploadFileController> _ilogger;

        public UploadFileController(ILogger<UploadFileController> logger)
        {
            _ilogger = logger;
        }

        // GET: api/<UploadFileController>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<UploadFileController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        // POST api/<UploadFileController>
        [HttpPost]
        public async Task<IActionResult> ReadExcelFrom_FormData(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("File is empty.");
            }

            var data = new List<Dictionary<string, string>>();

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
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

        // PUT api/<UploadFileController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<UploadFileController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}

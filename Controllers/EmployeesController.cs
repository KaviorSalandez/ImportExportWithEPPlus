using System.IO;
using DemoImportExport.Caches;
using DemoImportExport.Consts;
using DemoImportExport.DTOs.Employees;
using DemoImportExport.Helper;
using DemoImportExport.Models;
using DemoImportExport.Models.Response;
using DemoImportExport.Services.EmployeeServices;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace DemoImportExport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmployeesController : ControllerBase
    {
        private readonly IEmployeeService _employeeService;
        public ICacheService _cacheService;

        public EmployeesController(IEmployeeService employeeService, ICacheService cacheService)
        {
            _employeeService = employeeService;
            _cacheService = cacheService;
        }

        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            var employees = await _employeeService.GetAllAsync();
            return Ok(employees);
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetById(int id)
        {
            var employee = await _employeeService.GetByIdAsync(id);
            if (employee == null)
                return NotFound();

            return Ok(employee);
        }

        [HttpPost]
        public async Task<IActionResult> Create([FromBody] Employee employee)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            await _employeeService.AddAsync(employee);
            return CreatedAtAction(nameof(GetById), new { id = employee.EmployeeId }, employee);
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> Update(int id, [FromBody] Employee employee)
        {
            if (id != employee.EmployeeId)
                return BadRequest("ID mismatch");

            await _employeeService.UpdateAsync(employee);
            return NoContent();
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(int id)
        {
            await _employeeService.DeleteAsync(id);
            return NoContent();
        }


        [HttpGet]
        [Route("Filter")]

        public async Task<IActionResult> FindAllFilter(int pageSize = 0, int pageNumber = 1, string? search = "", string? email = "")
        {
            if (pageSize == 0)
            {
                var employees = await _employeeService.GetAllAsync();
                pageSize = employees.Count();
            }
            if (search == null)
            {
                search = "";
            }

            var entities = await _employeeService.FindAllFilter(pageSize, pageNumber, search, email);

            return StatusCode(200, entities);

        }


        // xuất file mẫu + xuất file data

        [HttpGet("ExportExcel")]
        public async Task<IActionResult> ExportExcel([FromQuery] bool isFileMau = true, [FromQuery] List<int> listID = null)
        {

            byte[] excelData = await _employeeService.ExportExcel(isFileMau, listID);
            string fileName = $"List_Employee_{DateTime.Now.ToString("dd/MM/yy")}.xlsx";
            return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        [HttpPost(RoutesConst.ImportEmployeeAPI)]
        public async Task<ActionResult<ApiResponse<object>>> ImportEmployee(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new ApiResponse<object>
                {
                    Status = 400,
                    Message = "File không hợp lệ.",
                    Data = null,
                    Error = "No file uploaded"
                });
            }

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0; // Reset the stream position to the beginning
                // Read the Excel file and map it to EmployeeExcelDto
                var listEmployee = await HelperFile.ReadExcel<EmployeeExcelDto>(stream);
                stream.Close();

                return Ok(new ApiResponse<object>
                {
                    Status = 200,
                    Message = "Import thành công",
                    Data = listEmployee
                });
            }
            catch (Exception ex)
            {
                return StatusCode(400, new ApiResponse<object>
                {
                    Status = 400,
                    Message = "Lỗi xử lý file",
                    Error = ex.Message
                });
            }
        }

        //[HttpGet("ExportExcelFail/{id}")]
        //public async Task<IActionResult> ExportExcelFail(string id)
        //{
        //    byte[] excelData = await _quizQuestionService.ExportExcel(1, id);
        //    string fileName = $"Quiz-Fail-{DateTime.Now.ToString("dd-MM-yy HH:mm:ss")}.xlsx";
        //    return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //}

        //[HttpGet("ExportExcelResult/{id}")]
        //public async Task<IActionResult> ExportExcelResult(string id)
        //{
        //    byte[] excelData = await _quizQuestionService.ExportExcel(1, id);
        //    string fileName = $"Quiz-Result-{DateTime.Now.ToString("dd-MM-yy HH:mm:ss")}.xlsx";
        //    return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //}

    }
}

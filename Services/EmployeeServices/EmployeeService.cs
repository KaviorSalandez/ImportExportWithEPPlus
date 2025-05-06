using ClosedXML.Excel;
using DemoImportExport.DTOs.Employees;
using DemoImportExport.Enums;
using DemoImportExport.Models;
using DemoImportExport.Repositories.EmployeeRepositories;
using MISA.AMISDemo.Core.DTOs.Employees;
using System.Globalization;
using System.Text.RegularExpressions;
using static DemoImportExport.Enums.CDKEnum;
using System.Data;
using System.Reflection;
using DemoImportExport.Repositories.DepartmentRepositories;
using DemoImportExport.Repositories.PositionRepositories;
using DemoImportExport.Caches;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using AutoMapper;
using DemoImportExport.DTOs.Employee;
using DocumentFormat.OpenXml.Bibliography;
using OfficeOpenXml.Style;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DemoImportExport.Extensions;

namespace DemoImportExport.Services.EmployeeServices
{
    public class EmployeeService : IEmployeeService
    {
        private readonly IEmployeeRepository _employeeRepo;
        private readonly IDepartmentRepository _departmentRepository;
        private readonly IPositionRepository _positionRepository;
        private readonly ICacheService _cacheService;
        private readonly IMapper _mapper;


        public EmployeeService(IEmployeeRepository employeeRepo, IDepartmentRepository departmentRepository, IPositionRepository positionRepository, ICacheService cacheService, IMapper mapper)
        {
            _employeeRepo = employeeRepo;
            _departmentRepository = departmentRepository;
            _positionRepository = positionRepository;
            _cacheService = cacheService;
            _mapper = mapper;
        }

        public async Task<IEnumerable<Employee>> GetAllAsync()
        {
            return await _employeeRepo.GetAllAsync();
        }

        public async Task<Employee?> GetByIdAsync(int id)
        {
            return await _employeeRepo.GetByIdAsync(id);
        }

        public async Task AddAsync(Employee employee)
        {
            await _employeeRepo.AddAsync(employee);
        }

        public async Task UpdateAsync(Employee employee)
        {
            await _employeeRepo.UpdateAsync(employee);
        }

        public async Task DeleteAsync(int id)
        {
            await _employeeRepo.DeleteAsync(id);
        }

        public async Task<EmployeeCountDto> FindAllFilter(int pageSize = 10, int pageNumber = 1, string search = "", string? email = "")
        {
            IEnumerable<Employee> employees = await _employeeRepo.FindAllFilter(pageSize, pageNumber, search, email);
            EmployeeCountDto result = new EmployeeCountDto();

            if (employees != null && employees.Any())
            {
                // Map entities to DTOs using AutoMapper
                var employeeDto = _mapper.Map<IEnumerable<EmployeeDto>>(employees);

                // Assuming TotalRecord is a property of EmployeeCountDto
                result.Count = employees.FirstOrDefault().TotalRecord;

                result.Employees = employeeDto.ToList();
            }
            else
            {
                result.Count = 0;
                result.Employees = new List<EmployeeDto>(); // Ensure Employees list is initialized
            }

            return result;
        }


        public async Task<byte[]> ExportExcel(bool isFileMau, List<int>? Ids = null)
        {
            IEnumerable<EmployeeExcelDto> data = new List<EmployeeExcelDto>();
            // check Data  = 1 -> kết xuất file rỗng để import
            if(isFileMau)
            {
                return GenerateExcelFile(data,null);
            }
            else
            {
                // check Data = 2 -> kết xuất file có dữ liệu
                if (Ids != null && Ids.Count > 0)
                {
                    // chỉ kết xuất các bản ghi được tick
                    var employees = await _employeeRepo.FindManyRecord(Ids);
                    data = _mapper.Map<IEnumerable<EmployeeExcelDto>>(employees);
                }
                else
                {
                    // kết xuất tất các các bản ghi trong DB
                    var employees = await _employeeRepo.GetAllAsync();
                    data = _mapper.Map<IEnumerable<EmployeeExcelDto>>(employees);

                }
                return GenerateExcelFile(data,null);
            }
        }
        // style for column 
        public void StyleColumn(string[] columnHeaders, ExcelWorksheet worksheet, int dataStartRow, int endRow, string? keyRedis)
        {
            int[] columnWidths = { 7, 20, 30, 15, 15, 30, 30, 20, 20 };

            if (keyRedis != null)
            {
                int[] extendedColumnWidths = new int[columnWidths.Length + 1];
                for (int i = 0; i < columnWidths.Length; i++)
                {
                    extendedColumnWidths[i] = columnWidths[i];
                }

                // Add the new column header
                extendedColumnWidths[columnWidths.Length] = 50;
                // Replace the old array with the new one
                columnWidths = extendedColumnWidths;
            }

            // Add column headers and set column widths
            for (int i = 0; i < columnHeaders.Length; i++)
            {
                worksheet.Cells[dataStartRow, i + 1].Value = columnHeaders[i];
                worksheet.Cells[dataStartRow, i + 1].Style.WrapText = true;
                worksheet.Column(i + 1).Width = columnWidths[i];
                worksheet.Row(dataStartRow).Style.Font.Size = 12; // Set font size for headers row
                worksheet.Row(dataStartRow).Style.Font.Bold = true;
                // Apply border style to the entire column
                using (var range = worksheet.Cells[dataStartRow, i + 1, endRow, i + 1])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
            }
        }

        // create combobox 
        private void AddDataValidation(ExcelWorksheet worksheet, string[] columnHeaders, int startRow, int endRow,
            IEnumerable<EGender> comboBoxValues)
        {
            int genderColumnIndex = Array.IndexOf(columnHeaders, "Giới tính");
            if (genderColumnIndex == -1)
            {
                throw new ArgumentException("Column 'Giới tính' not found in columnHeaders.");
            }

            var validationRange = worksheet.Cells[startRow, genderColumnIndex + 1, endRow, genderColumnIndex + 1];

            var validation = validationRange.DataValidation.AddListDataValidation();

            foreach (var value in comboBoxValues)
            {
                validation.Formula.Values.Add(value.GetDisplayNameEnum());
            }
        }


        private byte[] GenerateExcelFile(IEnumerable<EmployeeExcelDto> data, string keyRedis)
        {   
            EGender[] typeGenders = (EGender[])Enum.GetValues(typeof(EGender));

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Bắt buộc nếu dùng EPPlus 5+
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Danh sách nhân viên");
                // Merge and style title
                ws.Cells["A1:I2"].Merge = true;
                ws.Cells["A1"].Value = "DANH SÁCH NHÂN VIÊN";
                ws.Cells["A1"].Style.Font.Size = 25;
                ws.Cells["A1"].Style.Font.Bold = true;
                ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Row(1).Height = 25;

                // style for column 
                string[] columnHeaders = {
                    "STT",
                    "Mã Nhân Viên",
                    "Tên Nhân Viên",
                    "Giới tính",
                    "Ngày Sinh",
                    "Tên Vị trí",
                    "Tên Đơn Vị",
                    "Số Tài Khoản",
                    "Tên Ngân Hàng"
                };

                if (keyRedis != null)
                {
                    string[] extendedColumnHeaders = new string[columnHeaders.Length + 1];
                    for (int i = 0; i < columnHeaders.Length; i++)
                    {
                        extendedColumnHeaders[i] = columnHeaders[i];
                    }

                    // Add the new column header
                    extendedColumnHeaders[columnHeaders.Length] = "Tình trạng";
                    // Replace the old array with the new one
                    columnHeaders = extendedColumnHeaders;
                }

                int dataStartRow = 3;
                int endRow = 1000;
                StyleColumn(columnHeaders, ws, dataStartRow, endRow, keyRedis);


                int startRow = dataStartRow + 1; // Assuming data starts from row (dataStartRow + 1)
                AddDataValidation(ws, columnHeaders, startRow, endRow, typeGenders);

                ToConvertDataTable<EmployeeExcelDto>(data, ws);

                using (MemoryStream stream = new MemoryStream())
                {
                    package.SaveAs(stream);
                    return stream.ToArray();
                }
            }

        }

        /// <summary>
        /// kiểm tra các ids được truyền tới xem có rỗng và có trong database không 
        /// </summary>
        /// <param name="Ids">Danh sách các id </param>
        /// <returns></returns>
        /// <exception cref="ValidateException">trả về exception ko có id hoặc không tồn tại trong Database</exception>
        public async Task ValidateManyIds(List<int> Ids)
        {
            for (int i = 0; i < Ids.Count; i++)
            {
                var Id = Ids[i];
                if (Id == null)
                {
                    throw new Exception("ID không tìm thấy");
                }
            }

            for (int i = 0; i < Ids.Count; i++)
            {
                var Id = Ids[i];
                var find = await _employeeRepo.GetByIdAsync(Id);
                if (find == null)
                {
                    throw new Exception("Không tìm thấy đối tượng");
                }
            }
        }

        public Task<string> GenerateCode()
        {
            throw new NotImplementedException();
        }

        public async Task<EmployeeImportParentDto> ImportExcel(IFormFile fileImport)
        {
            await CheckFileImport(fileImport);

            int countSuccess = 0, countFail = 0;
            // Kết quả tổng hợp toàn bộ file
            var employeeImportParentDtos = new EmployeeImportParentDto();

            // list các bản ghi đã import (mỗi bản ghi có kiểm tra hợp lệ và ds lỗi)
            var employeeImportDtos = new List<EmployeeImportDto>();
            // Danh sách nhân viên hợp lệ (đã mapping)

            var employeeImportSuccess = new List<Employee>();

            var positions = await _positionRepository.GetAllAsync();
            var departments = await _departmentRepository.GetAllAsync();

            using (var stream = new MemoryStream())
            {
                // copy vào tệp stream 
                fileImport.CopyTo(stream);
                // thực hiện đọc dữ liệu trong file
                using (var package = new ExcelPackage(stream))
                {
                    // Đọc worksheet đầu 
                    ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
                    if (workSheet != null)
                    {
                        var rowCount = workSheet.Dimension.Rows;

                        for (int row = 4; row <= rowCount; row++)
                        {
                            var employeeImportDto = new EmployeeImportDto();
                            var gender = workSheet?.Cells[row, 4]?.Value?.ToString()?.Trim();
                            var dob = workSheet.Cells[row, 5]?.Value?.ToString()?.Trim();
                            var positionName = workSheet?.Cells[row, 6]?.Value?.ToString()?.Trim();
                            var departmentName = workSheet?.Cells[row, 7]?.Value?.ToString()?.Trim();

                            var checkPositionName = CheckCoincidence(positions, positionName, "PositionName");
                            var checkDepartmentName = CheckCoincidence(departments, departmentName, "DepartmentName");

                            employeeImportDto = new EmployeeImportDto
                            {
                                EmployeeCode = workSheet?.Cells[row, 2]?.Value?.ToString()?.Trim(),
                                EmployeeName = workSheet?.Cells[row, 3]?.Value?.ToString()?.Trim(),

                                Gender = ConvertGender(gender),

                                DOB = dob != "" && dob != null ? ProcessDate(dob) : null,

                                PositionId = checkPositionName != null ? (int) checkPositionName.GetType().GetProperty("PositionId")?.GetValue(checkPositionName, null): 0,
                                DepartmentId = checkDepartmentName != null ? (int)checkDepartmentName.GetType().GetProperty("DepartmentId")?.GetValue(checkDepartmentName, null):0,
                                
                                BankAccount = workSheet?.Cells[row, 8]?.Value?.ToString()?.Trim(),
                                BankName = workSheet?.Cells[row, 9]?.Value?.ToString()?.Trim(),
                            };
                            bool check = true;
                            if (checkDepartmentName == null)
                            {
                                AddImportError(employeeImportDto, $"{"Không tìm thấy đơn vị"}: {departmentName}");
                                check = false;
                            }
                            if (checkPositionName == null)
                            {
                                AddImportError(employeeImportDto, $"{"Không tìm thấy vị trí"}: {positionName}");
                                check = false;
                            }

                            var employee = _mapper.Map<Employee>(employeeImportDto);

                            var checkEmployeeCode = await _employeeRepo.CheckEmployeeCode(employeeImportDto.EmployeeCode);
                            if (checkEmployeeCode != null)
                            {
                                AddImportError(employeeImportDto, "Mã Nhân Viên đã được tạo");
                                check = false;
                            }
                            if (check == true)
                            {
                                countSuccess++;
                                employeeImportDto.IsImported = true;
                                employeeImportSuccess.Add(employee);
                            }
                            if (checkEmployeeCode != null || checkDepartmentName == null || checkPositionName == null)
                            {
                                countFail++;
                            }

                            employeeImportDtos.Add(employeeImportDto);

                        }
                    }
                    employeeImportParentDtos.CountSuccess = countSuccess;
                    employeeImportParentDtos.CountFail = countFail;
                    employeeImportParentDtos.EmployeeImportDtos = employeeImportDtos;

                    var cacheKey = $"excel-import-data-{Guid.NewGuid()}"; // Use a unique key
                    employeeImportParentDtos.IdImport = cacheKey;
                    DateTimeOffset expiryTime = DateTimeOffset.Now.AddDays(1);
                    _cacheService.SetData<string>(cacheKey, JsonConvert.SerializeObject(employeeImportSuccess), expiryTime);
                }
            }

            return employeeImportParentDtos;
        }

        public int ImportDatabase(string idImport)
        {
            if (idImport == null)
            {
                throw new Exception();
            }
            var dataImport = _cacheService.GetData<string>(idImport);
            var jArray = JsonConvert.DeserializeObject<JArray>(dataImport);
            var employees = jArray?.ToObject<List<Employee>>();

            var create = _employeeRepo.InsertMany(employees);
            return create;
        }

        /// <summary>
        /// Kiểm tra file import 
        /// </summary>
        /// <param name="fileImport">File được import </param>
        /// <exception cref="ValidateException"></exception>
        /// 
        public async Task CheckFileImport(IFormFile fileImport)
        {
            if (fileImport == null || fileImport.Length == 0)
            {
                // Ném ngoại lệ ValidateException với thông báo lỗi cụ thể
                throw new ArgumentException("File không hợp lệ, vui lòng chọn file để tải lên.", nameof(fileImport));
            }
            if (!Path.GetExtension(fileImport.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                // Ném ngoại lệ với thông báo lỗi về định dạng file
                throw new InvalidOperationException("Định dạng file không hợp lệ. Chỉ hỗ trợ file Excel (.xlsx).");
            }
        }


        /// <summary>
        /// Convert ngày tháng năm 
        /// </summary>
        /// <param name="input"> chuỗi ngày tháng năm </param>
        /// <returns></returns>
        public DateTime? ProcessDate(string input)
        {
            // Regex để kiểm tra định dạng yyyy
            string yearRegex = @"^\d{4}$";

            // Regex để kiểm tra định dạng dd/MM/yyyy
            string ddMmYyRegex = @"^\d{1,2}/\d{1,2}/\d{4}$";

            // Regex để kiểm tra định dạng MM/yyyy
            string mmYyRegex = @"^\d{1,2}/\d{4}$";

            // Kiểm tra input bằng Regex
            if (Regex.IsMatch(input, yearRegex))
            {
                // Trả về ngày đầu tiên của năm được cung cấp
                return DateTime.ParseExact($"01/01/{input}", "dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
            else if (Regex.IsMatch(input, ddMmYyRegex))
            {
                // Tách chuỗi thành các phần
                string[] parts = input.Split('/');

                // Bổ sung "0" nếu phần ngày chỉ có một ký tự
                if (parts[0].Length == 1)
                {
                    parts[0] = "0" + parts[0];
                }
                if (parts[1].Length == 1)
                {
                    parts[1] = "0" + parts[1];
                }

                // Ghép chuỗi lại và định dạng thành dd/MM/yyyy
                string formattedDate = string.Join("/", parts);

                // Trả về ngày được định dạng
                return DateTime.ParseExact(formattedDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            }
            else if (Regex.IsMatch(input, mmYyRegex))
            {
                string[] parts = input.Split('/');
                if (parts[0].Length == 1)
                {
                    input = "0" + input;
                }

                // Trả về ngày đầu tiên của tháng được cung cấp
                return DateTime.ParseExact($"01/{input}", "dd/MM/yyyy", CultureInfo.InvariantCulture);
            }

            // Nếu input không hợp lệ, trả về null
            return null;
        }


        /// <summary>
        /// Convert giới tính 
        /// </summary>
        /// <param name="gender">tên giới tính nhận được </param>
        /// <returns>Giới tính </returns>
        public Gender ConvertGender(string gender)
        {
            string male = "Nam";
            string female = "Nữ";
            string other = "Khác";

            if (gender != null && gender.ToLower().Equals(male))
            {
                return Gender.Nam;
            }
            if (gender != null && gender.ToLower().Equals(female))
            {
                return Gender.Nữ;
            }
            if (gender != null && gender.ToLower().Equals(other))
            {
                return Gender.Khác;
            }
            return Gender.Nam;
        }

        // Hàm hỗ trợ thêm lỗi import vào danh sách
        private void AddImportError(EmployeeImportDto dto, string error)
        {
            dto.Errors.Add(error);
            dto.IsImported = false;
        }

        /// <summary>
        /// Tìm name trong items xem có không 
        /// </summary>
        /// <param name="items">danh sách các items </param>
        /// <param name="name">tên của giá trị muốn tìm </param>
        /// <param name="nameCompare">thuộc tính trong items muốn so sánh</param>
        /// <returns></returns>
        public object? CheckCoincidence(IEnumerable<object> items, string name, string nameCompare)
        {
            if (items == null || name == null || nameCompare == null)
            {
                return null; // Handle null inputs gracefully
            }

            // Use case-insensitive comparison (optional)
            var find = items.FirstOrDefault(item =>
            {
                // Kiểm tra xem đối tượng có thuộc tính của nameCompare không
                var nameProperty = item.GetType().GetProperty(nameCompare);
                if (nameProperty == null)
                {
                    // Không tìm thấy thuộc tính "Name", trả về false
                    return false;
                }

                // Lấy giá trị của thuộc tính "Name" và so sánh với 'name'
                var itemName = nameProperty.GetValue(item) as string;
                return itemName != null && itemName.Equals(name, StringComparison.OrdinalIgnoreCase);
            });

            return find;
        }


       

        /// <summary>
        /// Chuyển đổi dữ liệu sang các bảng của excel 
        /// </summary>
        /// <typeparam name="T">kiểu thực thể T muốn chuyển đổi </typeparam>
        /// <param name="items">mảng các thực thể kiểu T </param>
        /// <returns>datatable</returns>
        public DataTable ToConvertDataTable<T>(IEnumerable<T> items, ExcelWorksheet ws)
        {
            DataTable dt = new DataTable(typeof(T).Name);
            PropertyInfo[] propInfo = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);

            // Thêm cột số thứ tự
            dt.Columns.Add("STT", typeof(int));
            foreach (PropertyInfo prop in propInfo)
            {
                dt.Columns.Add(prop.Name);
            }

            int ordinalNumber = 1;
            int rowIndex = 4; // dòng bắt đầu ghi dữ liệu
            foreach (T item in items)
            {
                // STT
                ws.Cells[rowIndex, 1].Value = ordinalNumber;

                for (int i = 0; i < propInfo.Length; i++)
                {
                    var propValue = propInfo[i].GetValue(item, null);
                    if (propValue != null)
                    {
                        Type propType = propInfo[i].PropertyType;

                        // Nếu là kiểu Nullable<>
                        if (propType.IsGenericType && propType.GetGenericTypeDefinition() == typeof(Nullable<>))
                        {
                            Type underlyingType = Nullable.GetUnderlyingType(propType);

                            if (underlyingType == typeof(DateTime))
                            {
                                DateTime dateTimeValue = (DateTime)propValue;
                                ws.Cells[rowIndex, i + 2].Value = dateTimeValue.ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                ws.Cells[rowIndex, i + 2].Value = propValue.ToString();
                            }
                        }
                        else
                        {
                            ws.Cells[rowIndex, i + 2].Value = propValue.ToString();
                        }
                    }
                    else
                    {
                        ws.Cells[rowIndex, i + 2].Value = ""; // giá trị mặc định nếu null
                    }
                }

                ordinalNumber++;
                rowIndex++;
            }

            return dt;
        }


        #region
        public Employee MapCreateDtoToEntity(EmployeeCreateDto entity)
        {
            var entityDto = _mapper.Map<Employee>(entity);
            return entityDto;
        }

        public EmployeeDto MapEntityToDto(Employee entity)
        {
            var entityDto = _mapper.Map<EmployeeDto>(entity);
            return entityDto;
        }

        public Employee MapUpdateDtoToEntity(EmployeeUpdateDto updateDto, Employee entity)
        {
            updateDto.EmployeeId = entity.EmployeeId;
            var entityDto = _mapper.Map(updateDto, entity);
            return entityDto;
        }
        #endregion


        public async Task CheckBeforeInsert(EmployeeCreateDto entity)
        {
            var department = await _departmentRepository.GetByIdAsync(entity.DepartmentId);
            if (department == null)
            {
                throw new Exception("Không tìm thấy đơn vị");
            }
            var position = await _positionRepository.GetByIdAsync(entity.PositionId);
            if (position == null)
            {
                throw new Exception("Không tìm thấy vị trí");
            }
            var employeeCode = await _employeeRepo.CheckEmployeeCode(entity.EmployeeCode.ToLower());
            if (employeeCode != null)
            {
                throw new Exception("Mã Nhân Viên đã được tạo");
            }
        }


        public  async Task CheckBeforeUpdate(EmployeeUpdateDto entity)
        {
            var department = await _departmentRepository.GetByIdAsync(entity.DepartmentId);
            if (department == null)
            {
                throw new Exception("Không tìm thấy đơn vị");
            }
            var position = await _positionRepository.GetByIdAsync(entity.PositionId);
            if (position == null)
            {
                throw new Exception("Không tìm thấy vị trí");
            }
        }

        public async Task<byte[]> ExportExcel2(bool isFileMau, string? keyRedis)
        {
            IEnumerable<EmployeeExcelDto> data = new List<EmployeeExcelDto>();
            //kiểm tra xem một ICollection
            if (keyRedis != null)
            {
                var dataImport = _cacheService.GetData<List<EmployeeImportDto>>(keyRedis);

                data = dataImport.Select(e => _mapper.Map<EmployeeExcelDto>(e)).ToList();
            }

            return await GenerateExcelFile(data, keyRedis);
        }
    }
}

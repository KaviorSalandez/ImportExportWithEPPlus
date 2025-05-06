using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DemoImportExport.Consts;
using DemoImportExport.Enums;

namespace DemoImportExport.DTOs.Employees
{
    public class EmployeeExcelDto
    {
        // Mã nhân viên
        public string EmployeeCode { get; set; }
        // Tên Nhân Viên 
        public string EmployeeName { get; set; }
        [Required(ErrorMessage = CDKConst.ERRMSG_DepartmentId)]


        // Giới tính
        public CDKEnum.Gender? Gender { get; set; }


        // Ngày sinh
        public DateTime? DOB { get; set; }

        // Tên chức vụ 
        public string PositionName { get; set; }

        [Required(ErrorMessage = CDKConst.ERRMSG_EmployeeCode)]
        // Tên phòng ban 
        public string DepartmentName { get; set; }

        [Required(ErrorMessage = CDKConst.ERRMSG_PositionId)]
        

        // Số tài khoản ngân hàng
        public string? BankAccount { get; set; }

        // Tên ngân hàng
        public string? BankName { get; set; }
        

    }
}

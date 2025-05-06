using DemoImportExport.Models;
using Microsoft.EntityFrameworkCore;

namespace DemoImportExport.Repositories.DepartmentRepositories
{
    public interface IDepartmentRepository
    {
        Task<IEnumerable<Department>> GetAllAsync();
        Task<Department?> GetByIdAsync(int id);
        Task AddAsync(Department department);
        Task UpdateAsync(Department department);
        Task DeleteAsync(int id);
        Task<Department> CheckDepartmentName(string departmentName);
    }
}

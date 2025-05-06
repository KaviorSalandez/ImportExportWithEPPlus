using System.Linq;
using DemoImportExport.Models;
using DemoImportExport.Persistents;
using EntityFramework.BulkExtensions;
using Microsoft.EntityFrameworkCore;

namespace DemoImportExport.Repositories.EmployeeRepositories
{
    public class EmployeeRepository : IEmployeeRepository
    {
        private readonly AppDbContext _context;

        public EmployeeRepository(AppDbContext context)
        {
            _context = context;
        }

        public async Task<IEnumerable<Employee>> GetAllAsync()
        {
            return await _context.Employees.ToListAsync();
        }

        public async Task<Employee?> GetByIdAsync(int id)
        {
            return await _context.Employees.FindAsync(id);
        }

        public async Task AddAsync(Employee employee)
        {
            await _context.Employees.AddAsync(employee);
            await _context.SaveChangesAsync();
        }

        public async Task UpdateAsync(Employee employee)
        {
            _context.Employees.Update(employee);
            await _context.SaveChangesAsync();
        }

        public async Task DeleteAsync(int id)
        {
            var employee = await _context.Employees.FindAsync(id);
            if (employee != null)
            {
                _context.Employees.Remove(employee);
                await _context.SaveChangesAsync();
            }
        }

        public async Task<Employee> CheckEmployeeCode(string employeeCode)
        {
            var entity = await _context.Employees.FirstOrDefaultAsync(e => e.EmployeeCode == employeeCode);
            return entity;
        }

        public async Task<Employee> CheckBankAccount(string bankAccount)
        {
            var entity = await _context.Employees.FirstOrDefaultAsync(e => e.BankAccount != null && e.BankAccount.Trim() == bankAccount.Trim());
            return entity;
        }

        public int InsertMany(List<Employee> employees)
        {
            if (employees == null || !employees.Any())
                return 0;

            _context.BulkInsert(employees);
            return employees.Count();
        }

        public async Task<IEnumerable<Employee>> FindAllFilter(int pageSize = 10, int pageNumber = 1, string search = "", string? email = "")
        {
            var query = _context.Employees.AsQueryable();

            // Nếu có email, lọc theo Email
            if (!string.IsNullOrEmpty(email))
            {
                query = query.Where(e => e.Email == email);
            }

            // Nếu có tìm kiếm theo tên
            if (!string.IsNullOrEmpty(search))
            {
                query = query.Where(e => e.EmployeeName.Contains(search));
            }

            // Phân trang
            query = query.Skip((pageNumber - 1) * pageSize).Take(pageSize);

            return await query.ToListAsync();
        }

        public async Task<IEnumerable<Employee>> FindManyRecord(List<int> Ids)
        {

            var employees = await _context.Employees
                .Where(e => Ids.Contains(e.EmployeeId))
                .Include(e => e.Department)
                .Include(e => e.Position)
                .ToListAsync();

            return employees;
        }
    }
}

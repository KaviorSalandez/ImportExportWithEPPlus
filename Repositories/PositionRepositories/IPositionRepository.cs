using DemoImportExport.Models;

namespace DemoImportExport.Repositories.PositionRepositories
{
    public interface IPositionRepository
    {
        Task<IEnumerable<Position>> GetAllAsync();
        Task<Position?> GetByIdAsync(int id);
        Task AddAsync(Position position);
        Task UpdateAsync(Position position);
        Task DeleteAsync(int id);

        Task<Position> CheckPositionName(string PositionName);
    }
}

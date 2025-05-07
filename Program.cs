
using DemoImportExport.Caches;
using DemoImportExport.Mapping;
using DemoImportExport.Persistents;
using DemoImportExport.Repositories.DepartmentRepositories;
using DemoImportExport.Repositories.EmployeeRepositories;
using DemoImportExport.Repositories.PositionRepositories;
using DemoImportExport.Services.DepartmentServices;
using DemoImportExport.Services.EmployeeServices;
using DemoImportExport.Services.PositionServices;
using DemoImportExport.Uow;
using Microsoft.EntityFrameworkCore;
using NLog;
using NLog.Web;

namespace DemoImportExport
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Early init of NLog to allow startup and exception logging, before host is built
            var logger = NLog.LogManager.Setup().LoadConfigurationFromAppSettings().GetCurrentClassLogger();
            logger.Debug("init main");
            try
            {
                var builder = WebApplication.CreateBuilder(args);

                // Add services to the container.

                builder.Services.AddControllers();
                // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
                builder.Services.AddEndpointsApiExplorer();
                builder.Services.AddSwaggerGen();

                builder.Services.AddDbContext<AppDbContext>(options =>
                options.UseSqlServer(builder.Configuration.GetConnectionString("MyCnn")));
                // Repository
                builder.Services.AddScoped(typeof(IUnitOfWork), typeof(UnitOfWork));
                // Services
                builder.Services.AddScoped<ICacheService, CacheService>();

                builder.Services.AddScoped<IEmployeeService, EmployeeService>();
                builder.Services.AddScoped<IDepartmentService, DepartmentService>();
                builder.Services.AddScoped<IPositionService, PositionService>();
                builder.Services.AddScoped<ICacheService, CacheService>();
                builder.Services.AddAutoMapper(typeof(EmployeeProfile));

                // NLog: Setup NLog for Dependency injection
                builder.Logging.ClearProviders();
                builder.Host.UseNLog();

                var app = builder.Build();

                // Configure the HTTP request pipeline.
                if (app.Environment.IsDevelopment())
                {
                    app.UseSwagger();
                    app.UseSwaggerUI();
                    app.UseHsts();
                }

                // use scope ensure connect to db 
                using (var scope = app.Services.CreateScope())
                {
                    var dbContext = scope.ServiceProvider.GetRequiredService<AppDbContext>();
                    try
                    {
                        if (dbContext.Database.CanConnect())
                        {
                            dbContext.Database.Migrate();
                        }
                        else
                        {
                            throw new Exception("Not found Database.");
                        }
                    }
                    catch (Exception ex)
                    {

                        throw new Exception(ex.Message);
                    }
                }

                app.UseHttpsRedirection();

                app.UseAuthorization();


                app.MapControllers();

                app.Run();
            }
            catch (Exception exception)
            {
                // NLog: catch setup errors
                logger.Error(exception, "Stopped program because of exception");
                throw;
            }
            finally
            {
                // Ensure to flush and stop internal timers/threads before application-exit (Avoid segmentation fault on Linux)
                NLog.LogManager.Shutdown();
            }
        }
    }
}

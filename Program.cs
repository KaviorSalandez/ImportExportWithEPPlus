
using DemoImportExport.Caches;
using DemoImportExport.Mapping;
using DemoImportExport.Persistents;
using DemoImportExport.Repositories.DepartmentRepositories;
using DemoImportExport.Repositories.EmployeeRepositories;
using DemoImportExport.Repositories.PositionRepositories;
using DemoImportExport.Services.DepartmentServices;
using DemoImportExport.Services.EmployeeServices;
using DemoImportExport.Services.PositionServices;
using Microsoft.EntityFrameworkCore;
using System;

namespace DemoImportExport
{
    public class Program
    {
        public static void Main(string[] args)
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
            builder.Services.AddScoped<IDepartmentRepository, DepartmentRepository>();
            builder.Services.AddScoped<IPositionRepository, PositionRepository>();
            builder.Services.AddScoped<IEmployeeRepository, EmployeeRepository>();
            // Services
            builder.Services.AddScoped<ICacheService, CacheService>();

            builder.Services.AddScoped<IEmployeeService, EmployeeService>();
            builder.Services.AddScoped<IDepartmentService, DepartmentService>();
            builder.Services.AddScoped<IPositionService, PositionService>();
            builder.Services.AddScoped<ICacheService,CacheService>();
            builder.Services.AddAutoMapper(typeof(EmployeeProfile));

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
                if (dbContext.Database.EnsureCreated())
                {
                    dbContext.Database.Migrate();
                } else
                {
                    throw new Exception("Database not created");
                }
            }

            app.UseHttpsRedirection();

            app.UseAuthorization();


            app.MapControllers();

            app.Run();
        }
    }
}

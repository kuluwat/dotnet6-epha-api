

using Model;
using Microsoft.Extensions.FileProviders;
using System.Reflection;
using Microsoft.OpenApi.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Builder;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddCors(p => p.AddPolicy("AllowOrigin", builder =>
{
    builder.WithOrigins("*").AllowAnyMethod().AllowAnyHeader();
}));


// Add services to the container.
builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    try
    {
        // using System.Reflection;
        var xmlFilename = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";

        c.IncludeXmlComments(Path.Combine(AppContext.BaseDirectory, xmlFilename));
    }
    catch (Exception ex) { }
});
builder.Services.AddDirectoryBrowser();
// เผื่อได้ใช้เนื่องจาก batch ไม่ได้ access ให้เขียนไฟล์
// builder.Services.AddHostedService<OwnerCronJob>();
builder.Services.AddControllers(options => { options.AllowEmptyInputInBodyModelBinding = true; })
.AddJsonOptions(opt =>
{
    opt.JsonSerializerOptions.PropertyNameCaseInsensitive = true;
    opt.JsonSerializerOptions.PropertyNamingPolicy = null;
});

#region Allow Policy Service SAP
var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
builder.Services.AddCors(options =>
{
    options.AddPolicy(MyAllowSpecificOrigins,
                          policy =>
                          {
                              policy.WithOrigins("*")
                                                  .AllowAnyHeader()
                                                  .AllowAnyMethod();

                          });
});
#endregion Allow Policy Service SAP

var app = builder.Build();
Config.setConfig(app.Services.GetRequiredService<IConfiguration>());
app.UseCors(MyAllowSpecificOrigins);
string logPath = app.Configuration["appsettings:folder_Logs"];
bool folderExists = Directory.Exists(logPath);
if (folderExists)
{
    app.UseFileServer(new FileServerOptions
    {
        FileProvider = new PhysicalFileProvider(
               Path.Combine(logPath, "folder_Log")),
        RequestPath = "/log",
        EnableDirectoryBrowsing = true
    });

    app.UseFileServer(new FileServerOptions
    {
        FileProvider = new PhysicalFileProvider(
               Path.Combine(logPath, "pic")),
        RequestPath = "/pic",
        EnableDirectoryBrowsing = true
    });
}

// Configure the HTTP request pipeline.
//if (app.Environment.IsDevelopment() || app.Environment.IsProduction())
//{
app.UseSwagger();
app.UseSwaggerUI();
//}


app.UseHttpsRedirection();
app.UseAuthorization();
app.UseStaticFiles();
app.MapControllers();
app.Run();



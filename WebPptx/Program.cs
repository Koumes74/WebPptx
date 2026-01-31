using Microsoft.OpenApi;
using WebPptx.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "WebPptx API",
        Version = "v1",
        Description = "API for extracting texts and attachments from PPTX files."
    });
});

builder.Services.AddScoped<IPptxRebuildService, PptxRebuildService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(options =>
    {
        options.SwaggerEndpoint("/swagger/v1/swagger.json", "WebPptx API v1");
    });
}

app.UseAuthorization();

app.MapControllers();

app.Run();

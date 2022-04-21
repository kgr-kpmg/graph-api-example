using Microsoft.Identity.Web;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
IConfiguration configuration = builder.Configuration;
builder.Services
    .AddMicrosoftIdentityWebApiAuthentication(configuration, "AzureAd")
    .EnableTokenAcquisitionToCallDownstreamApi()
    .AddMicrosoftGraph(configuration.GetSection("Graph"))
    .AddInMemoryTokenCaches();

builder.Services.AddControllers();

var app = builder.Build();

// Configure the HTTP request pipeline.

app.UseHttpsRedirection();

app.UseAuthentication();
app.UseRouting();
app.UseAuthorization();

app.MapControllers();

app.Run();

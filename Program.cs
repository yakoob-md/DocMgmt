using DocMgmt.Data;
using DocMgmt.Services;
using DocMgmt.Controllers; // Ensures the NativeDocumentParser namespace is accessible

namespace DocMgmt
{
	public class Program
	{
		public static void Main(string[] args)
		{
			var builder = WebApplication.CreateBuilder(args);

			// 1. Configure Kestrel server limits for Linux/Docker environments (100 MB max request body)
			builder.Services.Configure<Microsoft.AspNetCore.Server.Kestrel.Core.KestrelServerOptions>(options =>
			{
				options.Limits.MaxRequestBodySize = 104857600; // 100 MB in bytes
			});

			// 2. Configure HTTP form multipart boundary constraints for heavy file streams
			builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(options =>
			{
				options.ValueLengthLimit = 104857600;
				options.MultipartBodyLengthLimit = 104857600; // 100 MB in bytes
			});

			// Register MVC services.
			builder.Services.AddControllersWithViews();

			// Register the application database version manager service.
			builder.Services.AddScoped<DocumentUploadService>();

			// Register the 100% first-party native OpenXML parsing service engine.
			builder.Services.AddScoped<NativeDocumentParser>();

			// Register the SQL Server connection factory.
			builder.Services.AddSingleton<
				IDbConnectionFactory,
				SqlConnectionFactory>();

			var app = builder.Build();

			// Configure the HTTP request pipeline.
			if (!app.Environment.IsDevelopment())
			{
				app.UseExceptionHandler("/Home/Error");
				app.UseHsts();
			}

			app.UseHttpsRedirection();
			app.UseStaticFiles(); // Falls back safely across standard ASP.NET Core environments
			app.UseRouting();
			app.UseAuthorization();

			app.MapStaticAssets();

			app.MapControllerRoute(
				name: "default",
				pattern: "{controller=Home}/{action=Index}/{id?}")
				.WithStaticAssets();

			app.Run();
		}
	}
}

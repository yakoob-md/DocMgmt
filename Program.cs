using DocMgmt.Data;
using DocMgmt.Services;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Server.Kestrel.Core;

namespace DocMgmt
{
	public class Program
	{
		public static void Main(string[] args)
		{
			var builder = WebApplication.CreateBuilder(args);

			// Allow request bodies up to the 100 MB upload limit.
			builder.Services.Configure<KestrelServerOptions>(options =>
			{
				options.Limits.MaxRequestBodySize = FileValidator.MaxFileBytes;
			});

			builder.Services.Configure<FormOptions>(options =>
			{
				options.ValueLengthLimit = (int)FileValidator.MaxFileBytes;
				options.MultipartBodyLengthLimit = FileValidator.MaxFileBytes;
			});

			builder.Services.AddControllersWithViews();

			builder.Services.AddScoped<DocumentUploadService>();

			// Stateless, so one shared instance is enough.
			builder.Services.AddSingleton<NativeDocumentParser>();

			builder.Services.AddSingleton<IDbConnectionFactory, SqlConnectionFactory>();

			var app = builder.Build();

			if (!app.Environment.IsDevelopment())
			{
				app.UseExceptionHandler("/Home/Error");
				app.UseHsts();
			}

			app.UseHttpsRedirection();
			app.UseStaticFiles();
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
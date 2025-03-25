using ExcelDataReader;
using ExcelDocTransfer.Context;
using ExcelDocTransfer.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace ExcelDocTransfer.Controllers
{
	public class DataUploadController : Controller
	{
		IConfiguration _configuration;
		IWebHostEnvironment _hostingEnvironment;
		DataContext _context;
		IExcelDataReader _excelDataReader;

		public DataUploadController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment, DataContext context)//, IExcelDataReader excelDataReader)
		{
			_configuration = configuration;
			_hostingEnvironment = hostingEnvironment;
			_context = context;
			//_excelDataReader = excelDataReader;
		}

		[HttpGet]
		public async Task<IActionResult> Index()
		{
			var data = await _context.CustomerResponses.ToListAsync();
			return View(data);
		}

		#region CoPilot
		/*
		[HttpPost]
		public async Task<IActionResult> Index(IFormFile file)
		{
			if (file == null || file.Length == 0)
			{
				ViewBag.Error = "Please select a file";
				return View();
			}
			if (!Path.GetExtension(file.FileName).Equals(".xls") && !Path.GetExtension(file.FileName).Equals(".xlsx"))
			{
				ViewBag.Error = "Invalid file format";
				return View();
			}
			var filePath = Path.Combine(_hostingEnvironment.WebRootPath, "files", file.FileName);
			using (var stream = new FileStream(filePath, FileMode.Create))
			{
				await file.CopyToAsync(stream);
			}
			using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
			{
				// Auto-detect format, supports:
				//  - Binary Excel files (2.0-2003 format; *.xls)
				//  - OpenXml Excel files (2007 format; *.xlsx)
				using (var reader = ExcelReaderFactory.CreateReader(stream))
				{
					do
					{
						while (reader.Read())
						{
							var customerResponse = new CustomerResponse
							{
								CustomerName = reader.GetValue(0).ToString(),
								CustomerAddress = reader.GetValue(1).ToString(),
								CustomerCity = reader.GetValue(2).ToString(),
								Fees = Convert.ToDecimal(reader.GetValue(3)),
								VisitDate = Convert.ToDateTime(reader.GetValue(4))
							};
							_context.CustomerResponses.Add(customerResponse);
						}
					} while (reader.NextResult());
				}
			}
			await _context.SaveChangesAsync();
			ViewBag.Success = "Data uploaded successfully";
			return View();
		}  
		*/
		#endregion

		[HttpPost]
		public async Task<IActionResult> Index(IFormFile file)
		{
			#region Arayüzden seçilen dosyanın wwwRoot klasörü altında istediğimiz bir klasöre kopyalama işlemi yapıyoruz.
			if (file == null || file.Length == 0)
			{
				ViewBag.Error = "Please select a file";
				return RedirectToAction(nameof(Index));
			}
			if (!Path.GetExtension(file.FileName).Equals(".xls") && !Path.GetExtension(file.FileName).Equals(".xlsx"))
			{
				ViewBag.Error = "Invalid file format";
				return View();
			}
			var filePath = Path.Combine(_hostingEnvironment.WebRootPath, "CustomerFiles", file.FileName);

			if (!System.IO.File.Exists(filePath))
				Directory.CreateDirectory(filePath);

			string dataFileName = Path.GetFileName(file.FileName);
			string extension = Path.GetExtension(dataFileName);
			string[] allowedExtensions = { ".xls", ".xlsx" };
			if (!allowedExtensions.Contains(extension))
			{
				ViewBag.Error = "Invalid file format";
				return View();
			}

			string saveToPath = Path.Combine(filePath, dataFileName);

			using (var fileStream = new FileStream(saveToPath, FileMode.Create))
			{
				await file.CopyToAsync(fileStream);
			}

			#endregion
			return View();
		}
	}
}

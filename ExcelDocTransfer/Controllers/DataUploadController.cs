using ExcelDataReader;
using ExcelDocTransfer.Context;
using ExcelDocTransfer.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Data;
using Microsoft.Extensions.Logging;

namespace ExcelDocTransfer.Controllers
{
	public class DataUploadController : Controller
	{
		IConfiguration _configuration;
		IWebHostEnvironment _hostingEnvironment;
		DataContext _context;
		IExcelDataReader _excelDataReader;
		int hataOlanExcelSatirId;
		private readonly ILogger<DataUploadController> _logger;

		public DataUploadController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment, DataContext context, ILogger<DataUploadController> logger)//, IExcelDataReader excelDataReader)
		{
			_configuration = configuration;
			_hostingEnvironment = hostingEnvironment;
			_context = context;
			_logger = logger;
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
			try
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
				var filePath = Path.Combine(_hostingEnvironment.WebRootPath, "CustomerFiles");

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

				//Dosya ile ilgili işlemler bitene kadar başkasının erişimine kapat
				//using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))

				System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

				using (var stream = new FileStream(saveToPath, FileMode.Open))
				{
					//_excelDataReader = ExcelReaderFactory.CreateBinaryReader(stream);
					_excelDataReader = ExcelReaderFactory.CreateReader(stream);

					//Datatable -> Tek bir veri kümesi	(Model)
					//DataSet -> Birden fazla veri kümesi (Tablo) 
					DataSet ds = new DataSet();
					ds = _excelDataReader.AsDataSet();
					_excelDataReader.Close();

					if (ds != null && ds.Tables.Count > 0)
					{
						//foreach (DataTable dt in ds.Tables)
						//{
						//	for (int i = 0; i < dt.Rows.Count; i++)
						//	{
						//		var customerResponse = new CustomerResponse
						//		{
						//			CustomerName = dt.Rows[i][0].ToString(),
						//			CustomerAddress = dt.Rows[i][1].ToString(),
						//			CustomerCity = dt.Rows[i][2].ToString(),
						//			Fees = Convert.ToDecimal(dt.Rows[i][3]),
						//			VisitDate = Convert.ToDateTime(dt.Rows[i][4])
						//		};
						//		await _context.CustomerResponses.AddAsync(customerResponse);
						//	}
						//}
						DataTable dt = ds.Tables[0];
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							hataOlanExcelSatirId = i;
							string InvoiceNumber = dt.Rows[i][0].ToString().Trim();
							bool isInvoiceNumberExists = await _context.CustomerResponses.AnyAsync(x => x.InvoiceNumber == InvoiceNumber);

							if (isInvoiceNumberExists) {
								_logger.LogWarning("Excel dosyasında bulunan müşteri adı zaten mevcut. Müşteri Adı: {InvoiceNumber}", InvoiceNumber);
								continue; // Eğer müşteri adı zaten varsa, bu satırı atla
							}
						//	DateTime xyz;
						//	DateTime.TryParse(dt.Rows[i][4].ToString(), out xyz);

							var customerResponse = new CustomerResponse
							{
								CustomerName = dt.Rows[i][0].ToString().Trim(),
								CustomerAddress = dt.Rows[i][1].ToString().Trim(),
								CustomerCity = dt.Rows[i][2].ToString().Trim(),
								InvoiceNumber = dt.Rows[i][3].ToString().Trim(),
								//Fees = Convert.ToDecimal(dt.Rows[i][3]),
								Fees = decimal.TryParse(dt.Rows[i][4].ToString(), out var fees) ? fees : 0,
								//VisitDate = Convert.ToDateTime(dt.Rows[i][4])
								VisitDate =DateTime.TryParse(dt.Rows[i][5].ToString(), out var visitDate) ? visitDate : DateTime.MinValue
							};


							await _context.CustomerResponses.AddAsync(customerResponse);
							await _context.SaveChangesAsync();
						}
					}

				}

				return RedirectToAction(nameof(Index));
			}
			catch (Exception ex)
			{
				_logger.LogError(ex, "Excel işlemei/okuma sırasında hata oluştu. Satır {Row}", hataOlanExcelSatirId);
				return View("Error", new ErrorViewModel
				{
					//içini doldur
					RequestId = hataOlanExcelSatirId.ToString(),
					ErrorMessage = ex.Message,
					StackTrace = ex.StackTrace ?? "",
					ActionName = this.RouteData.Values["action"]?.ToString() ?? "Bilinmiyor",
					ControllerName = this.RouteData.Values["controller"]?.ToString() ?? "Bilinmiyor",
				});
			}
		}
	}
}

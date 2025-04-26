using System.ComponentModel.DataAnnotations;

namespace ExcelDocTransfer.Models
{
	public class CustomerResponse
	{
		public int Id { get; set; }
		public string CustomerName { get; set; }
		public string CustomerAddress { get; set; }
		public string CustomerCity { get; set; }

		public string InvoiceNumber { get; set; }

		public decimal Fees { get; set; }

		[DisplayFormat(DataFormatString ="{0:dd.MM.yyyy}",ApplyFormatInEditMode =true)]
		public DateTime VisitDate { get; set; }
	}
}

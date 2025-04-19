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
		public DateTime VisitDate { get; set; }
	}
}

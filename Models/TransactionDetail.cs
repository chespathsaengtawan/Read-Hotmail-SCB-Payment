namespace ReadMail.Models
{
    public class TransactionDetail
    {
        public string? Method { get; set; }
        public BankAccount? From { get; set; }
        public decimal AmountBaht { get; set; }
        public string? ToAccount { get; set; }
        public string? DateTime { get; set; }
    }
}
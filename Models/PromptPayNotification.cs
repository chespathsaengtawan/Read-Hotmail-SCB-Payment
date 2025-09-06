namespace ReadMail.Models
{
    public class PromptPayNotification
    {
        public string? Recipient { get; set; }
        public TransactionDetail? Transaction { get; set; }
        public string? Sender { get; set; }
        public string? Note { get; set; }
    }
}
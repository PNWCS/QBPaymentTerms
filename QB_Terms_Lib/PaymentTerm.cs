namespace QB_Terms_Lib
{
    public class PaymentTerm(string qbID, string qbRev, string name, int companyID)
    {
        public string QB_ID { get; set; } = qbID;
        public string QB_Rev { get; set; } = qbRev;
        public string Name { get; set; } = name;
        public int Company_ID { get; set; } = companyID;
    }
}
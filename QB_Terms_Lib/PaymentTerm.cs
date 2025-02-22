namespace QB_Terms_Lib
{
    public class PaymentTerm
    {
        public string QB_ID { get; set; }
        public string QB_Rev { get; set; }
        public string Name { get; set; }
        public int Company_ID { get; set; }

        public PaymentTerm(string qbID, string qbRev, string name, int companyID)
        {
            QB_ID = qbID;
            QB_Rev = qbRev;
            Name = name;
            Company_ID = companyID;
        }

    }
}
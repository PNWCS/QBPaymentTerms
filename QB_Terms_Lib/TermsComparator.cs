using Serilog;

namespace QB_Terms_Lib
{
    public class TermsComparator
    {
        public static List<PaymentTerm> CompareTerms(List<PaymentTerm> companyTerms)
        {
            Log.Information("TermsComparator Initialized.");

            // Read QB terms
            List<PaymentTerm> qbTerms = TermsReader.QueryAllTerms();

            // Convert QuickBooks and Company terms into dictionaries for quick lookup
            var qbTermDict = qbTerms.ToDictionary(t => t.Company_ID, t => t);
            var companyTermDict = companyTerms.ToDictionary(t => t.Company_ID, t => t);

            List<PaymentTerm> newTermsToAdd = new List<PaymentTerm>();

            // Iterate through company terms to compare with QB terms
            foreach (var companyTerm in companyTerms)
            {
                if (qbTermDict.TryGetValue(companyTerm.Company_ID, out var qbTerm))
                {
                    // Term exists in both, compare names
                    if (qbTerm.Name == companyTerm.Name)
                    {
                        companyTerm.Status = PaymentTermStatus.Unchanged;
                        Log.Information($"Term {companyTerm.Name} is Unchanged.");
                    }
                    else
                    {
                        companyTerm.Status = PaymentTermStatus.Different;
                        Log.Information($"Term {companyTerm.Name} is Different.");
                    }
                }
                else
                {
                    // Term does not exist in QB, queue for addition
                    newTermsToAdd.Add(companyTerm);
                }
            }

            // Check for terms that exist in QB but not in the company file
            foreach (var qbTerm in qbTerms)
            {
                if (!companyTermDict.ContainsKey(qbTerm.Company_ID))
                {
                    qbTerm.Status = PaymentTermStatus.Missing;
                    Log.Information($"Term {qbTerm.Name} is Missing.");
                }
            }

            // Call TermsAdder to add new terms to QuickBooks
            if (newTermsToAdd.Count > 0)
            {
                TermsAdder.AddTerms(newTermsToAdd);

                // Ensure `Added` terms are updated in `companyTermDict`
                foreach (var addedTerm in newTermsToAdd)
                {
                    if (companyTermDict.TryGetValue(addedTerm.Company_ID, out var companyTerm))
                    {
                        companyTerm.Status = addedTerm.Status;
                        Log.Information($"Term {companyTerm.Name} is Added.");
                    }
                }
            }

            // Merge `companyTermDict` with `qbTermDict` (removing duplicates)
            Dictionary<int, PaymentTerm> mergedTermsDict = new Dictionary<int, PaymentTerm>();

            foreach (var term in qbTermDict.Values)
                mergedTermsDict[term.Company_ID] = term; // Add all QB terms

            foreach (var term in companyTermDict.Values)
                mergedTermsDict[term.Company_ID] = term; // Overwrite with company terms

            // Convert merged dictionary back to a list
            Log.Information("TermsComparator Completed");
            return mergedTermsDict.Values.ToList();
        }
    }
}

using System.Diagnostics;
using Serilog;
using QB_Terms_Lib; // Adjust to your actual namespace for TermsComparator, PaymentTerm, etc.
using QBFC16Lib; // Needed for QuickBooks interaction
using static QB_Terms_Test.CommonMethods;

namespace QB_Terms_Test
{
    [Collection("Sequential Tests")]
    public class TermsComparatorTests
    {
        [Fact]
        public void CompareTerms_InMemoryScenario_And_Verify_Logs()
        {

            // 1️ Create 5 new PaymentTerms in memory
            const int START_ID = 100;
            var initialTerms = new List<PaymentTerm>();

            // 1) Ensure Serilog has released file access before deleting old logs
            EnsureLogFileClosed();
            DeleteOldLogFiles();
            ResetLogger();

            for (int i = 0; i < 5; i++)
            {
                string termName = "TestTerm_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                int companyId = START_ID + i;
                initialTerms.Add(new PaymentTerm(termName, companyId));
            }

            List<PaymentTerm> firstCompareResult = new List<PaymentTerm>();
            List<PaymentTerm> secondCompareResult = new List<PaymentTerm>();

            try
            {
                // 2️ First Compare: Expect all terms to be 'Added' in QuickBooks
                firstCompareResult = TermsComparator.CompareTerms(initialTerms);

                foreach (var term in firstCompareResult.Where(t => initialTerms.Any(x => x.Company_ID == t.Company_ID)))
                {
                    Assert.Equal(PaymentTermStatus.Added, term.Status);
                }

                // 3️ Modify Terms:
                var updatedTerms = new List<PaymentTerm>(initialTerms);
                var termToRemove = updatedTerms[0]; // Simulate "Missing"
                var termToRename = updatedTerms[1]; // Simulate "Different"
                updatedTerms.Remove(termToRemove);
                termToRename.Name += "_Modified";

                // 4️ Second Compare: Expect 'Missing' & 'Different' statuses
                secondCompareResult = TermsComparator.CompareTerms(updatedTerms);
                var secondCompareDict = secondCompareResult.ToDictionary(t => t.Company_ID);

                Assert.True(secondCompareDict.ContainsKey(termToRemove.Company_ID), "Missing term was not found in results.");
                Assert.Equal(PaymentTermStatus.Missing, secondCompareDict[termToRemove.Company_ID].Status);

                Assert.True(secondCompareDict.ContainsKey(termToRename.Company_ID), "Renamed term was not found in results.");
                Assert.Equal(PaymentTermStatus.Different, secondCompareDict[termToRename.Company_ID].Status);

                var unaffectedIds = updatedTerms
                    .Select(t => t.Company_ID)
                    .Except(new[] { termToRename.Company_ID })
                    .ToArray();

                foreach (var id in unaffectedIds)
                {
                    var term = secondCompareDict[id];
                    Assert.Equal(PaymentTermStatus.Unchanged, term.Status);
                }
            }
            finally
            {
                var allAddedTerms = firstCompareResult?.Where(t => !string.IsNullOrEmpty(t.QB_ID)).ToList();

                if (allAddedTerms != null && allAddedTerms.Count > 0)
                {
                    using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
                    {
                        foreach (var term in allAddedTerms)
                        {
                            DeleteStandardTerm(qbSession, term.QB_ID);
                        }
                    }
                }
            }

            // 5) Ensure logs are fully flushed before accessing them
            EnsureLogFileClosed();

            // 6) Verify new log file exists
            string logFile = GetLatestLogFile();
            EnsureLogFileExists(logFile);  // Ensure log file is fully written before reading

            // 7) Read the log file content
            string logContents = File.ReadAllText(logFile);

            // 8) Assert expected log messages exist
            Assert.Contains("TermsComparator Initialized", logContents);
            Assert.Contains("TermsComparator Completed", logContents);

            // 9) Verify that each term status change was logged properly
            foreach (var term in firstCompareResult)
            {
                string expectedLogMessage = $"Term {term.Name} is {term.Status}.";
                Assert.Contains(expectedLogMessage, logContents);
            }
            foreach (var term in secondCompareResult)
            {
                string expectedLogMessage = $"Term {term.Name} is {term.Status}.";
                Assert.Contains(expectedLogMessage, logContents);
            }
        }
        private void DeleteStandardTerm(QuickBooksSession qbSession, string listID)
        {
            IMsgSetRequest requestMsgSet = qbSession.CreateRequestSet();
            IListDel listDelRq = requestMsgSet.AppendListDelRq();
            listDelRq.ListDelType.SetValue(ENListDelType.ldtStandardTerms);
            listDelRq.ListID.SetValue(listID);

            IMsgSetResponse responseMsgSet = qbSession.SendRequest(requestMsgSet);
            WalkListDelResponse(responseMsgSet, listID);
        }

        private void WalkListDelResponse(IMsgSetResponse responseMsgSet, string listID)
        {
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null || responseList.Count == 0) return;

            IResponse response = responseList.GetAt(0);
            if (response.StatusCode == 0 && response.Detail != null)
            {
                Debug.WriteLine($"Successfully deleted Standard Term (ListID: {listID}).");
            }
            else
            {
                Debug.WriteLine($"Error Deleting Standard Term: {response.StatusMessage}");
            }
        }
    }
}

using System.Diagnostics;
using QB_Terms_Lib; // For PaymentTerm, TermsAdder, AppConfig, etc.
using QBFC16Lib;    // For QBSessionManager
using static QB_Terms_Test.CommonMethods;

namespace QB_Terms_Test
{
    [CollectionDefinition("Sequential Tests", DisableParallelization = true)]
    public class TermsAdderTests
    {
        [Fact]
        public void AddMultipleTerms_BatchAdd_And_Verify_Logs()
        {
            // 1) Ensure Serilog has released file access and delete old logs BEFORE initializing logging.
            EnsureLogFileClosed();
            DeleteOldLogFiles();
            ResetLogger();

            const int TERM_COUNT = 10;
            const int STARTING_COMPANY_ID = 100;
            List<PaymentTerm> termsToAdd = new List<PaymentTerm>();

            // 3) Build a list of random PaymentTerm objects.
            for (int i = 0; i < TERM_COUNT; i++)
            {
                string randomName = "BatchTerm_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                int companyID = STARTING_COMPANY_ID + i;
                termsToAdd.Add(new PaymentTerm(randomName, companyID));
            }

            // 4) Add terms to QuickBooks.
            TermsAdder.AddTerms(termsToAdd);

            // 5) Verify each term was added successfully.
            foreach (var term in termsToAdd)
            {
                Assert.Equal(PaymentTermStatus.Added, term.Status);
                Assert.False(string.IsNullOrEmpty(term.QB_ID),
                    $"Term {term.Name} has Status=Added but QB_ID was null or empty.");
            }

            // 6) Confirm terms exist in QuickBooks.
            using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
            {
                var allQBTerms = QueryAllTermsDirectly(qbSession);

                foreach (var term in termsToAdd)
                {
                    var matchingTerm = allQBTerms.FirstOrDefault(t => t.QB_ID == term.QB_ID);
                    Assert.NotNull(matchingTerm);
                    Assert.Equal(term.Name, matchingTerm.Name);
                    Assert.Equal(term.Company_ID, matchingTerm.Company_ID);
                }

                // 7) Cleanup: delete newly added terms.
                foreach (var term in termsToAdd.Where(t => t.Status == PaymentTermStatus.Added
                                                    && !string.IsNullOrEmpty(t.QB_ID)))
                {
                    DeleteStandardTerm(qbSession, term.QB_ID);
                }
            }

            // 8) Ensure logs are fully flushed before accessing them.
            EnsureLogFileClosed();

            // 9) Verify new log file exists.
            string logFile = GetLatestLogFile();
            EnsureLogFileExists(logFile);  // Wait for the log file to be fully written.

            // 10) Read the log file content.
            string logContents = File.ReadAllText(logFile);

            // 11) Assert expected log messages exist.
            Assert.Contains("TermsAdder Initialized", logContents);
            Assert.Contains("TermsAdder Completed", logContents);

            // 12) Verify that each added term was logged properly.
            foreach (var term in termsToAdd)
            {
                string expectedLogMessage = $"Added a new payment term {term.Name} to QB";
                Assert.Contains(expectedLogMessage, logContents);
            }
        }

        private List<PaymentTerm> QueryAllTermsDirectly(QuickBooksSession qbSession)
        {
            IMsgSetRequest requestMsgSet = qbSession.CreateRequestSet();
            IStandardTermsQuery standardTermsQueryRq = requestMsgSet.AppendStandardTermsQueryRq();

            IMsgSetResponse responseMsgSet = qbSession.SendRequest(requestMsgSet);
            return ParseQueryResponse(responseMsgSet);
        }

        private List<PaymentTerm> ParseQueryResponse(IMsgSetResponse responseMsgSet)
        {
            List<PaymentTerm> results = new List<PaymentTerm>();
            if (responseMsgSet == null) return results;

            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return results;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                    if (responseType == ENResponseType.rtStandardTermsQueryRs)
                    {
                        IStandardTermsRetList standardTermsRetList = (IStandardTermsRetList)response.Detail;
                        results.AddRange(ExtractTermsFromResponse(standardTermsRetList));
                    }
                }
            }

            return results;
        }

        private List<PaymentTerm> ExtractTermsFromResponse(IStandardTermsRetList standardTermsRetList)
        {
            List<PaymentTerm> terms = new List<PaymentTerm>();
            for (int i = 0; i < standardTermsRetList.Count; i++)
            {
                var standardTerm = standardTermsRetList.GetAt(i);
                if (standardTerm == null) continue;

                string qbID = standardTerm.ListID?.GetValue() ?? "";
                string name = standardTerm.Name?.GetValue() ?? "";
                int companyID = standardTerm.StdDiscountDays?.GetValue() ?? -1;

                terms.Add(new PaymentTerm(qbID, name, companyID));
            }
            return terms;
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
                throw new Exception($"Error Deleting Standard Term (ListID: {listID}): {response.StatusMessage}. Status code: {response.StatusCode}");
            }
        }
    }
}

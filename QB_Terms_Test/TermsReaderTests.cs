using System.Diagnostics;
using Serilog;
using QB_Terms_Lib; // For PaymentTerm, TermsReader, AppConfig, etc.
using QBFC16Lib;    // For QuickBooks session and API interaction
using static QB_Terms_Test.CommonMethods;

namespace QB_Terms_Test
{
    [Collection("Sequential Tests")]
    public class TermsReaderTests
    {
        [Fact]
        public void AddAndReadMultipleTerms_FromQuickBooks_And_Verify_Logs()
        {
            const int TERM_COUNT = 5;
            const int STARTING_COMPANY_ID = 100;
            List<PaymentTerm> termsToAdd = new List<PaymentTerm>();

            // 1) Ensure Serilog has released file access before deleting old logs
            EnsureLogFileClosed();
            DeleteOldLogFiles();
            ResetLogger();

            // 2) Build a list of random PaymentTerm objects
            for (int i = 0; i < TERM_COUNT; i++)
            {
                string randomName = "ReadTestTerm_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                int companyID = STARTING_COMPANY_ID + i; // Ensure IDs are continuous but below 999
                termsToAdd.Add(new PaymentTerm(randomName, companyID));
            }

            // 3) Add terms directly to QuickBooks
            using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
            {
                foreach (var term in termsToAdd)
                {
                    string qbID = AddStandardTerm(qbSession, term.Name, term.Company_ID);
                    term.QB_ID = qbID; // Store the returned QB ID
                }
            }

            // 4) Query QuickBooks to retrieve all terms
            var allQBTerms = TermsReader.QueryAllTerms();

            // 5) Verify that all added terms are present in QuickBooks
            foreach (var term in termsToAdd)
            {
                var matchingTerm = allQBTerms.FirstOrDefault(t => t.QB_ID == term.QB_ID);
                Assert.NotNull(matchingTerm);
                Assert.Equal(term.Name, matchingTerm.Name);
                Assert.Equal(term.Company_ID, matchingTerm.Company_ID);
            }

            // 6) Cleanup: Delete the added terms
            using (var qbSession = new QuickBooksSession(AppConfig.QB_APP_NAME))
            {
                foreach (var term in termsToAdd.Where(t => !string.IsNullOrEmpty(t.QB_ID)))
                {
                    DeleteStandardTerm(qbSession, term.QB_ID);
                }
            }

            // 7) Ensure logs are fully flushed before accessing them
            EnsureLogFileClosed();

            // 8) Verify new log file exists
            string logFile = GetLatestLogFile();
            EnsureLogFileExists(logFile);  // Ensure log file is fully written before reading

            // 9) Read the log file content
            string logContents = File.ReadAllText(logFile);

            // 10) Assert expected log messages exist
            Assert.Contains("TermsReader Initialized", logContents);
            Assert.Contains("TermsReader Completed", logContents);

            // 11) Verify that each retrieved term was logged properly
            foreach (var term in termsToAdd)
            {
                string expectedLogMessage = $"Successfully retrieved {term.Name} from QB";
                Assert.Contains(expectedLogMessage, logContents);
            }
        }

        

        private string AddStandardTerm(QuickBooksSession qbSession, string name, int companyID)
        {
            IMsgSetRequest requestMsgSet = qbSession.CreateRequestSet();
            IStandardTermsAdd stdTermsAddRq = requestMsgSet.AppendStandardTermsAddRq();
            stdTermsAddRq.Name.SetValue(name);
            stdTermsAddRq.StdDiscountDays.SetValue(companyID);

            IMsgSetResponse responseMsgSet = qbSession.SendRequest(requestMsgSet);
            return ExtractListIDFromResponse(responseMsgSet);
        }

        private string ExtractListIDFromResponse(IMsgSetResponse responseMsgSet)
        {
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null || responseList.Count == 0)
                throw new Exception("No response from StandardTermsAddRq.");

            IResponse response = responseList.GetAt(0);
            if (response.StatusCode != 0)
                throw new Exception($"StandardTermsAdd failed: {response.StatusMessage}");

            IStandardTermsRet? stdTermRet = response.Detail as IStandardTermsRet;
            if (stdTermRet == null)
                throw new Exception("No IStandardTermsRet returned after adding Standard Term.");

            return stdTermRet.ListID?.GetValue() ?? throw new Exception("ListID is missing in QuickBooks response.");
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

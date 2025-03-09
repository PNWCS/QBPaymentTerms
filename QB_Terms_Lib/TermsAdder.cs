using QBFC16Lib;
using Serilog;

namespace QB_Terms_Lib
{
    public class TermsAdder
    {
        public static void AddTerms(List<PaymentTerm> terms)
        {
            Log.Information("TermsAdder Initialized.");
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager? sessionManager = null;

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", AppConfig.QB_APP_NAME);
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Send the request and get the response from QuickBooks
                foreach (var term in terms)
                {
                    // Create a new request set for each term
                    IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                    requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                    BuildStandardTermsAddRq(requestMsgSet, term);
                    IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                    WalkStandardTermsAddRs(responseMsgSet, term);
                }

                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                Log.Information("TermsAdder Completed.");
            }
            catch (Exception e)
            {
                Log.Error(e.Message);
                if (sessionBegun && sessionManager != null)
                {
                    sessionManager.EndSession();
                }
                if (connectionOpen && sessionManager != null)
                {
                    sessionManager.CloseConnection();
                }
            }
        }

        static void BuildStandardTermsAddRq(IMsgSetRequest requestMsgSet, PaymentTerm term)
        {
            IStandardTermsAdd StandardTermsAddRq = requestMsgSet.AppendStandardTermsAddRq();
            //Set field value for Name
            StandardTermsAddRq.Name.SetValue(term.Name);
            //Set field value for StdDiscountDays
            StandardTermsAddRq.StdDiscountDays.SetValue(term.Company_ID);
        }

        static void WalkStandardTermsAddRs(IMsgSetResponse responseMsgSet, PaymentTerm term)
        {
            if (responseMsgSet == null) return;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;
            Log.Debug($"responseList.Count = {responseList.Count}");

            //if we sent only one request, there is only one response, we'll walk the list for this sample
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                Log.Debug($"response.StatusCode = {response.StatusCode}");
                //check the status code of the response, 0=ok, >0 is warning
                if (response.StatusCode == 0)
                {
                    //the request-specific response is in the details, make sure we have some
                    if (response.Detail != null)
                    {
                        //make sure the response is the type we're expecting
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtStandardTermsAddRs)
                        {
                            //upcast to more specific type here, this is safe because we checked with response.Type check above
                            IStandardTermsRet StandardTermsRet = (IStandardTermsRet)response.Detail;
                            WalkStandardTermsRet(StandardTermsRet, term);
                        }
                    }
                }
                else
                {
                    term.Status = PaymentTermStatus.FailedToAdd;
                    Log.Error($"QB API Returned Error Code {response.StatusCode}");
                }
            }
        }

        static void WalkStandardTermsRet(IStandardTermsRet StandardTermsRet, PaymentTerm term)
        {
            if (StandardTermsRet == null) return;
            //Go through all the elements of IStandardTermsRet
            //Get value of Name
            string qbName = (string)StandardTermsRet.Name.GetValue();

            //Get value of StdDiscountDays
            int qbCompanyID = -1; // -1 means no companyID
            if (StandardTermsRet.StdDiscountDays != null)
            {
                qbCompanyID = (int)StandardTermsRet.StdDiscountDays.GetValue();
            }

            // (New) Extract ListID
            if (StandardTermsRet.ListID != null)
            {
                term.QB_ID = StandardTermsRet.ListID.GetValue();
            }

            if (qbName == term.Name && qbCompanyID == term.Company_ID)
            {
                term.Status = PaymentTermStatus.Added;
            }
            Log.Information($"Added a new payment term {term.Name} to QB");
        }
    }
}

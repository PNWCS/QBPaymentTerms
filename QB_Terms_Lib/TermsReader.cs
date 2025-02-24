

using QBFC16Lib;

namespace QB_Terms_Lib
{
    public class TermsReader
    {
        public static List<PaymentTerm> QueryAllTerms()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;
            List<PaymentTerm> terms = new List<PaymentTerm>();

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                IStandardTermsQuery StandardTermsQueryRq = requestMsgSet.AppendStandardTermsQueryRq();

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                terms = WalkStandardTermsQueryRs(responseMsgSet);
                return terms;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                if (sessionBegun)
                {
                    sessionManager.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager.CloseConnection();
                }
                return terms;
            }
        }
        //void BuildStandardTermsQueryRq(IMsgSetRequest requestMsgSet)
        //{
        //    IStandardTermsQuery StandardTermsQueryRq = requestMsgSet.AppendStandardTermsQueryRq();

        //}




        static List<PaymentTerm> WalkStandardTermsQueryRs(IMsgSetResponse responseMsgSet)
        {

            List<PaymentTerm> terms = new List<PaymentTerm>();
            if (responseMsgSet == null) return terms;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return terms;
            //if we sent only one request, there is only one response, we'll walk the list for this sample
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                //check the status code of the response, 0=ok, >0 is warning
                if (response.StatusCode >= 0)
                {
                    //the request-specific response is in the details, make sure we have some
                    if (response.Detail != null)
                    {
                        //make sure the response is the type we're expecting
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtStandardTermsQueryRs)
                        {
                            //upcast to more specific type here, this is safe because we checked with response.Type check above
                            IStandardTermsRetList StandardTermsRet = (IStandardTermsRetList)response.Detail;
                            terms = WalkStandardTermsRet(StandardTermsRet);
                        }
                    }
                }
            }
            return terms;
        }




        static List<PaymentTerm> WalkStandardTermsRet(IStandardTermsRetList StandardTermsRet)
        {

            List<PaymentTerm> terms = new List<PaymentTerm>();
            if (StandardTermsRet == null) return terms;
            //Go through all the elements of IStandardTermsRetList
            //Get value of ListID

            for (int i = 0; i < StandardTermsRet.Count; i++)
            {

                var term = StandardTermsRet.GetAt(i);
                string qbID = (string)term.ListID.GetValue();
                //Get value of EditSequence
                string qbRev = (string)term.EditSequence.GetValue();
                //Get value of Name
                string name = (string)term.Name.GetValue();
                int companyID = 0;
                //Get value of StdDiscountDays
                if (term.StdDiscountDays != null)
                {
                    companyID = (int)term.StdDiscountDays.GetValue();
                }

                Console.WriteLine($"{name}, {qbID} , {qbRev}, {companyID}");

                PaymentTerm paymentTerm = new PaymentTerm(qbID, qbRev, name, companyID);

                terms.Add(paymentTerm);

            }

            return terms;

        }

    }
}
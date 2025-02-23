using QBFC16Lib;
using System;
using System.Collections.Generic;

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
                // Create the session Manager object
                sessionManager = new QBSessionManager();

                // Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                IStandardTermsQuery StandardTermsQueryRq = requestMsgSet.AppendStandardTermsQueryRq();

                // Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                // Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                // End the session and close the connection to QuickBooks
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

        static List<PaymentTerm> WalkStandardTermsQueryRs(IMsgSetResponse responseMsgSet)
        {
            List<PaymentTerm> terms = new List<PaymentTerm>();
            if (responseMsgSet == null) return terms;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return terms;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode >= 0)
                {
                    if (response.Detail != null)
                    {
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtStandardTermsQueryRs)
                        {
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

            for (int i = 0; i < StandardTermsRet.Count; i++)
            {
                var term = StandardTermsRet.GetAt(i);
                string qbID = term.ListID.GetValue();
                string qbRev = term.EditSequence.GetValue();
                string name = term.Name.GetValue();
                int companyID = term.StdDiscountDays != null ? term.StdDiscountDays.GetValue() : 0;

                Console.WriteLine($"{name}, {qbID}, {qbRev}, {companyID}");

                PaymentTerm paymentTerm = new PaymentTerm(qbID, qbRev, name, companyID);
                terms.Add(paymentTerm);
            }
            return terms;
        }
    }
}

using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.SharePoint.Client;
//using WebServices.AlertsWebServiceReference;
using System.Web;
using System.Web.Services;
using System.Web.Services.Description;
using System.Web.Services.Protocols;

using Utils = MNIT.Utilities;

namespace MNIT.Inventory

    
{
    class GetAlerts
    {
        [SoapDocumentMethod("http://schemas.microsoft.com/sharepoint/soap/2002/1/alerts/GetAlerts", RequestNamespace = "http://schemas.microsoft.com/sharepoint/soap/2002/1/alerts/", ResponseNamespace = "http://schemas.microsoft.com/sharepoint/soap/2002/1/alerts/", Use = SoapBindingUse.Literal, ParameterStyle = SoapParameterStyle.Wrapped)] 

        public static void InventoryAlerts(string siteAddress, Utils.ActingUser actingUser)
        {
            
            ClientContext ctx = new ClientContext(siteAddress);
            ctx.Credentials = !string.IsNullOrEmpty(actingUser.UserLoginName) ? new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword, actingUser.UserDomain) : System.Net.CredentialCache.DefaultCredentials;
            Web subWeb = ctx.Web;


            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(siteAddress);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //// HTTP GET
                //HttpResponseMessage response = await client.GetAsync("api/products/1");
                //if (response.IsSuccessStatusCode)
                //{
                //    Product product = await response.Content.ReadAsAsync<Product>();
                //    Console.WriteLine("{0}\t${1}\t{2}", product.Name, product.Price, product.Category);
                //}
            }



            //Web_Reference_Folder_Name.Alerts alertService = new Web_Reference_Folder_Name.Alerts();
            //alertService.Credentials = System.Net.CredentialCache.DefaultCredentials;

            //Web_Reference_Folder_Name.AlertInfo allAlerts = alertService.GetAlerts();

            //Console.WriteLine("Server: " + allAlerts.AlertServerName +
            //    "\nURL: " + allAlerts.AlertServerUrl +
            //    "\nWeb Title: " + allAlerts.AlertWebTitle +
            //    "\nNumber: " + allAlerts.Alerts.Length.ToString());
            
            //GetAlerts alerts = new GetAlerts();
            //Console.WriteLine(alerts);
            
            //Alerts alerts = new Alerts();
            //alerts.Url = siteAddress + "/_vti_bin/alerts.asmx";
            ////alerts.Credentials = CredentialCache.DefaultCredentials;
            //alerts.Credentials = !string.IsNullOrEmpty(actingUser.UserLoginName) ? new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword, actingUser.UserDomain) : System.Net.CredentialCache.DefaultCredentials;
            //AlertInfo alertInfo = alerts.GetAlerts();
            //Console.WriteLine("AlertWebTitle: "+alertInfo.AlertWebTitle);
            //Console.WriteLine("AlertServerName: "+alertInfo.AlertServerName);
            //Console.WriteLine("AlertServerType: " + alertInfo.AlertServerType);
            //Console.WriteLine("AlertServerUrl: "+alertInfo.AlertServerUrl);
            //Console.WriteLine("Alerts Number:" + alertInfo.Alerts.Length.ToString());
            //Console.WriteLine("CurrentUser: "+alertInfo.CurrentUser);
            //foreach (Alert alert in alertInfo.Alerts)
            //{
            //    Console.WriteLine("Alert Information: ");
            //    Console.WriteLine("-------------------");
            //    Console.WriteLine("Title: "+alert.Title);
            //    Console.WriteLine("AlertForUrl: "+alert.AlertForUrl);
            //}

        }
    }
}

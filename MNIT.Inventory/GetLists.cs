using System;
using System.Linq;
using System.Net;
using Microsoft.SharePoint.Client;

namespace MNIT.Inventory
{
    public class GetLists
    {
        // Method to inventory information about sites with workflows and instances of workflows
        public static void InventoryLists(string siteAddress, Utilities.ActingUser actingUser, ref int largeListCounter, ref int unlimitedVerCounter, ref int infoPathFormCounter, ref int infoPathExternalConnCounter, string csvFilePath)
        {
            ClientContext ctx = new ClientContext(siteAddress);
            ctx.Credentials = !string.IsNullOrEmpty(actingUser.UserLoginName) ? new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword, actingUser.UserDomain) : CredentialCache.DefaultCredentials;
            Web subWeb = ctx.Web;
            // Load web and web properties
            ctx.Load(ctx.Web, w => w.Webs, w => w.Url, w => w.Title, w => w.Lists, w => w.Id);
            // Execute Query against web
            ctx.ExecuteQuery();
            // Initialize variables
            string currentWebTitle = ctx.Web.Title;
            string currentWebUrl = ctx.Web.Url;
            Uri tempUri = new Uri(currentWebUrl);
            string urlDomain = tempUri.Host;
            string urlProtocol = tempUri.Scheme;
            //int versionCount = 0;
            string unlimitedVersions = null;
            string siteCollId = null;
            string webId = null;
            // find the SCAs or owners of the site collection
            ctx.Load(ctx.Site, sc => sc.Owner, sc => sc.RootWeb, sc => sc.Id);
            ctx.ExecuteQuery();
            string rootWebOwner = ctx.Site.Owner.Email;
            if (string.IsNullOrEmpty(rootWebOwner))
            {
                rootWebOwner = ctx.Site.Owner.Title;
            }
            // Only get the web ID and the Site Collection Web ID if it is not an App web
            if (currentWebUrl.ElementAt(8) != 'a')
            {
                siteCollId = ctx.Site.RootWeb.Id.ToString();
                webId = ctx.Web.Id.ToString();
            }

            foreach (List tmpList in ctx.Web.Lists)
            {
                // Initialize variables
                string strLargeListCount = null;
                string currentListTitle = "";
                // Load list and list properties
                ctx.Load(tmpList, t => t.Title, t => t.DefaultViewUrl, t => t.ItemCount, t => t.EnableVersioning, t => t.IsPrivate, t => t.Hidden, t => t.MajorVersionLimit, t => t.MajorWithMinorVersionsLimit);
                // Execute Query against the list
                ctx.ExecuteQuery();

                string currentListUrl = urlProtocol + "://" + urlDomain + tmpList.DefaultViewUrl;
                // Add the list to the stream if it has 5000 list items or more
                if (tmpList.ItemCount > 4999)
                {
                    largeListCounter++;
                    strLargeListCount = tmpList.ItemCount.ToString();
                    currentListTitle = tmpList.Title;
                }

                if (tmpList.Hidden != true && tmpList.IsPrivate != true && tmpList.EnableVersioning == true && (tmpList.MajorVersionLimit == 0 || tmpList.MajorWithMinorVersionsLimit == 0))
                {
                    unlimitedVerCounter++;
                    unlimitedVersions = "Unlimited Versions";
                    currentListTitle = tmpList.Title;
                }
                // If the list item count is over 4999 write a line to the stream that calls out a large list potential issue
                if (!string.IsNullOrEmpty(currentListTitle))
                {
                    // Write the information about large lists to the inventory CSV file
                    //WriteToStream(siteCollId, webId, currentWebTitle, currentWebUrl, rootWebOwner, currentListTitle, currentListUrl, strLargeListCount, unlimitedVersions, null, null, streamWriter);
                    string[] passingListObject = new string[10];
                    passingListObject[0] = csvFilePath;
                    passingListObject[1] = siteCollId;
                    passingListObject[2] = webId;
                    passingListObject[3] = currentWebTitle;
                    passingListObject[4] = currentWebUrl;
                    passingListObject[5] = rootWebOwner;
                    passingListObject[6] = currentListTitle;
                    passingListObject[7] = currentListUrl;
                    passingListObject[8] = strLargeListCount;
                    passingListObject[9] = unlimitedVersions;
                    //passingListObject[10] = infoPathForm;
                    //passingListObject[11] = externalConnections;
                    WriteReports.WriteText(passingListObject);
                }
            }
            // Recursively inventory lists; Only use webs and sub web that are not a host for apps
            foreach (var recursiveSubWeb in ctx.Web.Webs)
            {
                if (recursiveSubWeb.Url.ElementAt(8) != 'a')
                {
                    InventoryLists(recursiveSubWeb.Url, actingUser, ref largeListCounter, ref unlimitedVerCounter, ref infoPathFormCounter, ref infoPathExternalConnCounter, csvFilePath);
                }
            }
        }
    }
}

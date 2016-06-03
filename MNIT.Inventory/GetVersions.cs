using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

using Utils = MNIT.Utilities;
namespace MNIT.Inventory
{
    public class GetVersions
    {
        // Method to inventory information about sites with workflows and instances of workflows
        public static void InventoryVersions(string siteAddress, Utilities.ActingUser actingUser, ref int largeListCounter, ref int unlimitedVerCounter, ref int siteCollCheckedOut, string csvFilePath)
        {
            ClientContext ctx = new ClientContext(siteAddress);
            try
            {
                if (string.IsNullOrEmpty(actingUser.UserLoginName))
                {
                    ctx.Credentials = CredentialCache.DefaultCredentials;
                }
                else
                {
                    if (actingUser.UserLoginName.IndexOf("@") != -1)
                    {
                        ctx.Credentials = new SharePointOnlineCredentials(actingUser.UserLoginName,
                            actingUser.UserPassword);
                    }
                    else
                    {
                        ctx.Credentials = new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword,
                            actingUser.UserDomain);
                    }
                }
                //ctx.Credentials = !string.IsNullOrEmpty(actingUser.UserLoginName) ? new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword, actingUser.UserDomain) : CredentialCache.DefaultCredentials;
            
                Web subWeb = ctx.Web;
                // Load web and web properties
                ctx.Load(subWeb, w => w.Webs, w => w.Url, w => w.Title, w => w.Lists, w => w.Id);
                // Execute Query against web
                ctx.ExecuteQuery();
                // Initialize variables
                string currentWebTitle = subWeb.Title;
                string currentWebUrl = subWeb.Url;
                Uri tempUri = new Uri(currentWebUrl);
                string urlDomain = tempUri.Host;
                string urlProtocol = tempUri.Scheme;
                //int versionCount = 0;
                string unlimitedVersions = "";
                string siteCollId = "";
                string webId = null;
                // find the SCAs or owners of the site collection
                Site siteCollection = ctx.Site;
                ctx.Load(siteCollection, sc => sc.Owner, sc => sc.RootWeb, sc => sc.Id);
                ctx.ExecuteQuery();
                string rootWebOwner = siteCollection.Owner.Email;
                if (string.IsNullOrEmpty(rootWebOwner))
                {
                    rootWebOwner = siteCollection.Owner.Title;
                }
                // Only get the web ID and the Site Collection Web ID if it is not an App web
                if (currentWebUrl.ElementAt(8) != 'a')
                {
                    //siteCollId = siteCollection.RootWeb.Id.ToString();
                    siteCollId = siteCollection.Id.ToString();
                    webId = subWeb.Id.ToString();
                }

                foreach (List tmpList in subWeb.Lists)
                {
                    // Initialize variables
                    string strTotalListItemCount = "";
                    string currentListTitle = "";
                    // Load list and list properties
                    ctx.Load(tmpList, t => t.Title, t => t.DefaultViewUrl, t => t.ItemCount, t => t.EnableVersioning,
                        t => t.EnableMinorVersions, t => t.IsPrivate, t => t.Hidden, t => t.MajorVersionLimit,
                        t => t.MajorWithMinorVersionsLimit, t => t.BaseType, t => t.ForceCheckout, t => t.Id);
                    // Execute Query against the list
                    ctx.ExecuteQuery();
                    // Build the URL
                    string currentListUrl = urlProtocol + "://" + urlDomain + tmpList.DefaultViewUrl;
                    // Build the Web Application Name
                    string webApplication = urlDomain.Split('.')[0];
                    // Add the list to the stream if it has 5000 list items or more
                    // Run the function to query large lists needs to be moved here to further optimize
                    if (tmpList.ItemCount > 4999)
                    {
                        largeListCounter++;
                        strTotalListItemCount = tmpList.ItemCount.ToString();
                        currentListTitle = tmpList.Title;
                    }

                    // Unlimited Versioning Check
                    if (tmpList.Hidden != true && tmpList.IsPrivate != true && tmpList.EnableVersioning == true &&
                        tmpList.EnableMinorVersions == true)
                    {

                        if ((tmpList.MajorVersionLimit == 0 || tmpList.MajorVersionLimit > 10) ||
                            (tmpList.MajorWithMinorVersionsLimit == 0 || tmpList.MajorWithMinorVersionsLimit > 10))
                        {
                            unlimitedVerCounter++;
                            unlimitedVersions = String.Format("Major:{0};Minor:{1}", tmpList.MajorVersionLimit,
                                tmpList.MajorWithMinorVersionsLimit);
                            currentListTitle = tmpList.Title;
                        }
                    }

                    if (tmpList.Hidden != true && tmpList.IsPrivate != true && tmpList.EnableVersioning == true &&
                        tmpList.EnableMinorVersions != true)
                    {
                        if (tmpList.MajorVersionLimit == 0 || tmpList.MajorVersionLimit > 10)
                        {
                            unlimitedVerCounter++;
                            unlimitedVersions = String.Format("Major:{0}", tmpList.MajorVersionLimit);
                            currentListTitle = tmpList.Title;
                        }
                    }

                    // counter including folders
                    int totalListItemCount = tmpList.ItemCount;
                    // Get item count and compare to checked in document count
                    int folderCount = 0;
                    // counter excluding folders
                    int listItemCount = 0;
                    int checkedInCount = 0;
                    int checkedOutCount = 0;
                    int neverCheckedInCount = 0;
                    string strFileCount = "";
                    string strFolderCount = "";
                    string strCheckedInCount = "";
                    string strCheckedOutCount = "";
                    string strNeverCheckedInCount = "";
                    string manageListUrl = "";
                    int largeListDiv = 3;

                    if (tmpList.BaseType == BaseType.DocumentLibrary)
                    {
                        if (totalListItemCount < 5000)
                        {
                            // Get a count of folders to be removed from the total list item count for comparing to checked in docs
                            var folders = tmpList.GetItems(CreateAllFoldersQuery());
                            ctx.Load(folders, icol => icol.Include(i => i.File));
                            ctx.ExecuteQuery();
                            foreach (var folder in folders)
                            {
                                File fileFolder = folder.File;
                                ctx.Load(fileFolder);
                                ctx.ExecuteQuery();
                                folderCount++;
                            }
                            // Get the files from the list
                            var items = tmpList.GetItems(CreateAllFilesQuery());
                            ctx.Load(items, icol => icol.Include(i => i.File, i => i.DisplayName));
                            ctx.ExecuteQuery();
                            foreach (var listItem in items)
                            {
                                File file = listItem.File;
                                ctx.Load(file, f => f.CheckOutType);
                                ctx.ExecuteQuery();
                                listItemCount++;
                                if (file.CheckOutType.ToString() == "None")
                                {
                                    checkedInCount++;
                                    //if (listItem["Thumbnail"].ToString().Length > 0)
                                    //{
                                    //}
                                    //else
                                    //{
                                    //    checkedInCount++;
                                    //}
                                }
                                else
                                {
                                    checkedOutCount++;
                                }
                            }
                            // Calculate the list item count without the folders
                            listItemCount = Math.Abs(totalListItemCount - folderCount);
                            // Add the list title so it gets included in the detailed report
                            if (checkedInCount > 0 && checkedInCount != listItemCount)
                            {
                                currentListTitle = tmpList.Title;
                                // prepare the total list item count
                                strTotalListItemCount = totalListItemCount.ToString();
                                // prepare the non folder list item count
                                strFileCount = listItemCount.ToString();
                                // prepare the folder count
                                strFolderCount = folderCount.ToString();
                                // prepare the Checked In count
                                strCheckedInCount = checkedInCount.ToString();
                                // prepare the Checked Out count
                                strCheckedOutCount = checkedOutCount.ToString();
                                // prepare the Never been checked in count
                                neverCheckedInCount = Math.Abs(listItemCount - checkedInCount - checkedOutCount);
                                //strNeverCheckedInCount = neverCheckedInCount.ToString();

                                FieldCollection fieldColl = tmpList.Fields;
                                ctx.Load(fieldColl);
                                ctx.ExecuteQuery();

                                foreach (Field fieldTemp in fieldColl)
                                {
                                    if (fieldTemp.InternalName == "AlternateThumbnailUrl")
                                    {
                                        strNeverCheckedInCount = fieldTemp.StaticName;
                                    }
                                    else
                                    {
                                        strNeverCheckedInCount = neverCheckedInCount.ToString();
                                    }
                                }
                                if (neverCheckedInCount > 0)
                                {
                                    manageListUrl = subWeb.Url + "/_layouts/15/ManageCheckedOutFiles.aspx?List={" +
                                                    tmpList.Id + "}";
                                }
                                // add to the site collection checked out counter
                                siteCollCheckedOut += checkedOutCount + neverCheckedInCount;
                            }
                        }
                        else
                        {
                            InventoryListItems(siteAddress, currentListTitle, 1000);
                        }



                    }
                    if (!string.IsNullOrEmpty(currentListTitle))
                    {
                        // Write the information about large lists to the inventory CSV file
                        string[] passingListObject = new string[17];
                        passingListObject[0] = csvFilePath;
                        passingListObject[1] = webApplication;
                        passingListObject[2] = siteCollId;
                        passingListObject[3] = webId;
                        passingListObject[4] = currentWebTitle;
                        passingListObject[5] = currentWebUrl;
                        passingListObject[6] = rootWebOwner;
                        passingListObject[7] = currentListTitle;
                        passingListObject[8] = currentListUrl;
                        passingListObject[9] = unlimitedVersions;
                        passingListObject[10] = strTotalListItemCount;
                        passingListObject[11] = strFolderCount;
                        passingListObject[12] = strFileCount;
                        passingListObject[13] = strCheckedInCount;
                        passingListObject[14] = strCheckedOutCount;
                        passingListObject[15] = strNeverCheckedInCount;
                        passingListObject[16] = manageListUrl;
                        WriteReports.WriteText(passingListObject);
                    }
                }

                // Recursively inventory lists; Only use webs and sub web that are not a host for apps
                foreach (var recursiveSubWeb in subWeb.Webs)
                {
                    if (recursiveSubWeb.Url.ElementAt(8) != 'a')
                    {
                        InventoryVersions(recursiveSubWeb.Url, actingUser, ref largeListCounter, ref unlimitedVerCounter, ref siteCollCheckedOut,
                            csvFilePath);
                    }
                }
            }
            catch (Exception ex23Exception)
            {
                Console.WriteLine("Could not gather all checked out document or list version information in site {0}.  {1}", siteAddress, ex23Exception);
            }
            finally
            {
                ctx.Dispose();
            }
        }

        public static CamlQuery CreateAllFilesQuery()
        {
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";
            //qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query><RowLimit></RowLimit></View>";
            return qry;
        }


        public static void InventoryListItems(string siteAddress, string listTitle, int rowLimit)
        {
            ClientContext ctx = new ClientContext(siteAddress);
            List list = ctx.Web.Lists.GetByTitle(listTitle); 
            ctx.Load(list);
            ctx.ExecuteQuery();
            ListItemCollectionPosition itemPosition = null;
            int largeCheckedInCount = 0;
            do
            {
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ListItemCollectionPosition = itemPosition;

                string viewXml = string.Format(@"
                        <View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <Eq>
                                        <FieldRef Name='FSObjType' />
                                        <Value Type='Integer'>0</Value>
                                    </Eq>
                                </Where>
                            </Query>
                            <ViewFields>
                                <FieldRef Name='Title' />
                            </ViewFields>
                            <RowLimit>{0}</RowLimit>
                        </View>", rowLimit);
                //camlQuery.ViewXml = "<View>" + Constants.vbCr + Constants.vbLf + "<ViewFields>" + Constants.vbCr + Constants.vbLf + "<FieldRef Name='Id'/><FieldRef Name='Title'/><FieldRef Name='Serial_No'/><FieldRef Name='CRM_ID'/>" + Constants.vbCr + Constants.vbLf + "</ViewFields>" + Constants.vbCr + Constants.vbLf + "<RowLimit>2201</RowLimit>" + Constants.vbCr + Constants.vbLf + "</View>";
                camlQuery.ViewXml = viewXml;
                ListItemCollection listItems = list.GetItems(camlQuery);
                ctx.Load(listItems);
                ctx.ExecuteQuery();
                itemPosition = listItems.ListItemCollectionPosition;
                foreach (ListItem listItem in listItems)
                {
                    //ctx.Load(listItem, li => li.DisplayName);
                    //ctx.ExecuteQuery();
                    File file = listItem.File;
                    ctx.Load(file, f => f.CheckOutType, f => f.Name);
                    ctx.ExecuteQuery();
                    //listItemCount++;
                    if (file.CheckOutType.ToString() == "None")
                    {
                        //checkedInCount++;
                        largeCheckedInCount++;
                    }
                    else
                    {
                        //checkedOutCount++;
                    }
                    //Console.WriteLine("Item Title: {0} Checked Out Type: {1}", listItem["Title"], file.CheckOutType);
                }
            }
            while (itemPosition != null);
            //Console.WriteLine(itemPosition.PagingInfo);
            Utils.SpinAnimation.Stop();
            Console.WriteLine();
            Console.WriteLine("Could not work with " + listTitle + ", because it is a large list.");
            Utils.SpinAnimation.Start();
        }

        private async Task<List<ListItem>> GetListItems(string siteAddress, string listTitle, int rowLimit)
        {
            List<ListItem> items = new List<ListItem>();
            ClientContext ctx = new ClientContext(siteAddress);
            List list = ctx.Web.Lists.GetByTitle(listTitle);
            //int rowLimit = 100;
            ListItemCollectionPosition position = null;

            string viewXml = string.Format(@"
                        <View Scope='RecursiveAll'>
                            <Query>
                                <Where>
                                    <Eq>
                                        <FieldRef Name='FSObjType' />
                                        <Value Type='Integer'>0</Value>
                                    </Eq>
                                </Where>
                            </Query>
                            <ViewFields>
                                <FieldRef Name='Title' />
                            </ViewFields>
                            <RowLimit>{0}</RowLimit>
                        </View>", rowLimit);

            var camlQuery = new CamlQuery();
            camlQuery.ViewXml = viewXml;

            do
            {
                ListItemCollection listItems = null;
                if (listItems != null && listItems.ListItemCollectionPosition != null)
                {
                    camlQuery.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                }

                listItems = list.GetItems(camlQuery);
                ctx.Load(listItems);
                //Task contextTask = ctx.ExecuteQueryAsync();
                //await Task.WhenAll(contextTask);

                //Task contextTask = ctx.ExecuteQuery();
                ctx.ExecuteQuery();
                await Task.WhenAll();
                position = listItems.ListItemCollectionPosition;
                items.AddRange(listItems.ToList());
            }
            while (position != null);
            return items;
        }

//        private async Task<List<ListItem>> GetListItems(string siteAddress, string listTitle)
//        {
//            List<ListItem> items = new List<ListItem>();
//            ClientContext ctx = new ClientContext(siteAddress);
//            //using (ClientContext context = SharePointContext.GetSharePointContext())
//            //{
//            List list = ctx.Web.Lists.GetByTitle(listTitle);
//            int rowLimit = 100;
//            ListItemCollectionPosition position = null;

//            string viewXml = string.Format(@"
//                <View Scope='RecursiveAll'>
//                    <Query>
//                        <Where>
//                            <Eq>
//                                <FieldRef Name='FSObjType' />
//                                <Value Type='Integer'>0</Value>
//                            </Eq>
//                        </Where>
//                    </Query>
//                    <ViewFields>
//                        <FieldRef Name='Title' />
//                    </ViewFields>
//                    <RowLimit>{0}</RowLimit>
//                </View>", rowLimit);

//            var camlQuery = new CamlQuery();
//            camlQuery.ViewXml = viewXml;

//            do
//            {
//                ListItemCollection listItems = null;
//                if (listItems != null && listItems.ListItemCollectionPosition != null)
//                {
//                    camlQuery.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
//                }

//                listItems = list.GetItems(CamlQuery.CreateAllItemsQuery());
//                ctx.Load(listItems);
//                //Task contextTask = ctx.ExecuteQueryAsync();
//                Task contextTask = ctx.ExecuteQuery();
//                await Task.WhenAll(contextTask);
//                position = listItems.ListItemCollectionPosition;
//                items.AddRange(listItems.ToList());
//            }
//            while (position != null);
//            //}
//            return items;
//        }


        public static void InventoryCheckedOut()
        {
            
        } 

        public static void DivideLargeList(int itemCount, int divisor)
        {
            Utils.SpinAnimation.Stop();
            Console.WriteLine();
            Console.WriteLine(@"Large List item count: {0}; Divisor: {1}", itemCount, divisor);
            Utils.SpinAnimation.Start();
            int remainderComparison = 300;
            // Get the Quotient
            int dividedCount = itemCount / divisor;
            // If the Quotient is larger than 5000
            if (dividedCount >= remainderComparison)
            {
                divisor++;
                DivideLargeList(itemCount, divisor);
            }
            else
            {
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Large List Dividend count: {0}; Divisor: {1}; Quotient: {2}", itemCount, divisor, dividedCount);
                Utils.SpinAnimation.Start();
            }
        }

        public static CamlQuery CreateAllFoldersQuery()
        {
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"ContentType\" /><Value Type=\"Text\">Folder</Value></Eq></Where></Query></View>";
            return qry;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using File = Microsoft.SharePoint.Client.File;
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

                    if (totalListItemCount < 5000)
                    {
                        // Check regular list configuration for problematic settings
                        if (tmpList.BaseType == BaseType.DocumentLibrary)
                        {
                            // Get a count of folders to be removed from the total list item count for comparing to checked in docs
                            var folders = tmpList.GetItems(CreateAllFoldersQuery());
                            ctx.Load(folders, icol => icol.Include(i => i.File));
                            ctx.ExecuteQuery();
                            folderCount = folders.Count;

                            // Get the checked in files from the list
                            var itemsCheckedIn = tmpList.GetItems(CreateCheckedInFilesQuery());
                            ctx.Load(itemsCheckedIn, icol => icol.Include(i => i.File, i => i.DisplayName));
                            ctx.ExecuteQuery();
                            checkedInCount = itemsCheckedIn.Count;


                            // Get the checked out files from the list
                            var itemsCheckedOut = tmpList.GetItems(CreateCheckedOutFilesQuery());
                            ctx.Load(itemsCheckedOut, icol => icol.Include(i => i.File, i => i.DisplayName));
                            ctx.ExecuteQuery();
                            checkedOutCount = itemsCheckedOut.Count;

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
                    }
                    else
                    {
                        // Check large list configuration for problematic settings
                        largeListCounter++;
                        totalListItemCount = tmpList.ItemCount;
                        strTotalListItemCount = totalListItemCount.ToString();
                        currentListTitle = tmpList.Title;

                        string qryType = "";
                        // Gather information about large lists
                        qryType = "folders";
                        folderCount = InventoryLargeLists(siteAddress, currentListTitle, qryType, 200, actingUser);
                        // checked in items
                        qryType = "checkedin";
                        checkedInCount = InventoryLargeLists(siteAddress, currentListTitle, qryType, 200, actingUser);
                        // checked out items
                        qryType = "checkedout";
                        checkedOutCount = InventoryLargeLists(siteAddress, currentListTitle, qryType, 200, actingUser);

                        // Calculate the list item count without the folders
                        listItemCount = Math.Abs(totalListItemCount - folderCount);
                        // Add the list title so it gets included in the detailed report
                        if (checkedInCount > 0 && checkedInCount != listItemCount)
                        {
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

                            strNeverCheckedInCount = neverCheckedInCount.ToString();
                            if (neverCheckedInCount > 0)
                            {
                                manageListUrl = subWeb.Url + "/_layouts/15/ManageCheckedOutFiles.aspx?List={" +
                                                tmpList.Id + "}";
                            }
                            // add to the site collection checked out counter
                            siteCollCheckedOut += checkedOutCount + neverCheckedInCount;
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

        public static int InventoryLargeLists(string siteAddress, string listTitle, string qryType, int rowLimit, Utils.ActingUser actingUser)
        {
            ClientContext ctx = new ClientContext(siteAddress);
            Web currentWeb = ctx.Web;
            int maxListItems = 5000;
            int lowerId = -1;
            int upperId = maxListItems;
            List<string> list = new List<string>();

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
            List tmpList = currentWeb.Lists.GetByTitle(listTitle);
            ctx.Load(tmpList);
            ctx.ExecuteQuery();

            int itemCount = 0;
            int runs = 0;
            int paginations = 1;
            if (tmpList.BaseType == BaseType.DocumentLibrary)
            {
                // get the number of times the query should be run against this library
                int neededRuns = (tmpList.ItemCount / 5000)+1;
                //Console.WriteLine("Needed runs according to max list limit " + neededRuns);
                // Query for folders in the large list
                if (qryType == "folders")
                {
                    // 0 out the itemPosition
                    ListItemCollectionPosition itemPosition = null;
                    for (int nr = 0; nr < neededRuns; nr++)
                    {
                        do
                        {
                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ListItemCollectionPosition = itemPosition;

                            // Get a count of folders to be removed from the total list item count for comparing to checked in docs
                            var listObjects = tmpList.GetItems(CreateLargeFoldersQuery(lowerId, upperId, rowLimit));
                            ctx.Load(listObjects, icol => icol.ListItemCollectionPosition,
                                icol => icol.Include(i => i.File));
                            ctx.ExecuteQuery();
                            itemCount += listObjects.Count;

                            itemPosition = listObjects.ListItemCollectionPosition;
                            lowerId += maxListItems;
                            upperId += maxListItems;
                            //Console.WriteLine("start row{0}  end row{1}", lowerId, upperId);
                            paginations++;
                        } while (itemPosition != null);
                        runs++;
                    }
                }
                // Query for checked in docs in the large list
                if (qryType == "checkedin")
                {
                    // 0 out the itemPosition
                    ListItemCollectionPosition itemPosition = null;
                    for (int nr = 0; nr < neededRuns; nr++)
                    {
                        do
                        {
                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ListItemCollectionPosition = itemPosition;
                            // Get a count of checked in docs
                            var listObjects =
                                tmpList.GetItems(CreateLargeCheckedInFilesQuery(lowerId, upperId, rowLimit));
                            ctx.Load(listObjects, icol => icol.ListItemCollectionPosition,
                                icol => icol.Include(i => i.File));
                            ctx.ExecuteQuery();
                            itemCount += listObjects.Count;

                            itemPosition = listObjects.ListItemCollectionPosition;
                            lowerId += maxListItems;
                            upperId += maxListItems;
                            //Console.WriteLine("start row{0}  end row{1}", lowerId, upperId);
                            paginations++;
                        } while (itemPosition != null);
                        runs++;
                    }
                }
                // Query for checked out docs in the large list
                if (qryType == "checkedout")
                {
                    // 0 out the itemPosition
                    ListItemCollectionPosition itemPosition = null;
                    for (int nr = 0; nr < neededRuns; nr++)
                    {
                        do
                        {
                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ListItemCollectionPosition = itemPosition;
                            // Get a count of checked out docs
                            var listObjects =
                                tmpList.GetItems(CreateLargeCheckedOutFilesQuery(lowerId, upperId, rowLimit));
                            ctx.Load(listObjects, icol => icol.ListItemCollectionPosition,
                                icol => icol.Include(i => i.File));
                            ctx.ExecuteQuery();
                            itemCount += listObjects.Count;

                            itemPosition = listObjects.ListItemCollectionPosition;
                            lowerId += maxListItems;
                            upperId += maxListItems;
                            //Console.WriteLine("start row{0}  end row{1}", lowerId, upperId);
                            paginations++;
                        } while (itemPosition != null);
                        runs++;
                    }
                }
            }
            //Console.WriteLine("{0} list has {1} {2} objects, paginated {3}, ran {4} times", listTitle, itemCount, qryType, paginations, runs);
            return itemCount;
        }


        public static CamlQuery CreateAllFoldersQuery()
        {
            // Create the list query for Folders
            // have used this <Eq><FieldRef Name='ContentType' /><Value Type='Text'>Folder</Value></Eq>
            // but this seems to be more effective <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq>
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope='RecursiveAll'><Query><Where>" +
                          "<Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq>" +
                          "</Where></Query></View>";
            return qry;
        }

        //public static CamlQuery CreateAllFilesQuery()
        //{
        //    var qry = new CamlQuery();
        //    qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";
        //    return qry;
        //}

        public static CamlQuery CreateCheckedInFilesQuery()
        {
            // Create the list query for checked IN files
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope='RecursiveAll'><Query><Where>" +
                              "<And>" +
                                "<IsNull><FieldRef Name='CheckoutUser' LookupId='TRUE' /></IsNull>" +
                                "<Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq>" +
                              "</And>" +
                          "</Where></Query></View>";
            return qry;
        }

        public static CamlQuery CreateCheckedOutFilesQuery()
        {
            // Create the list query for checked OUT files
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope='RecursiveAll'><Query><Where>" +
                              "<And>" +
                                "<IsNotNull><FieldRef Name='CheckoutUser' LookupId='TRUE' /></IsNotNull>" +
                                "<Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq>" +
                              "</And>" +
                          "</Where></Query></View>";
            return qry;
        }

        public static CamlQuery CreateLargeFoldersQuery(int lowerId, int upperId, int rowLimit)
        {
            // Create the large list query for Folders
            // <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq>
            // <Eq><FieldRef Name='ContentType'></FieldRef><Value Type='Text'>Folder</Value></Eq>
            var qry = new CamlQuery();
            qry.ViewXml = string.Format(@"
                        <View Scope='RecursiveAll'>
                            <Query>
                                <Where>
	                              <And>
		                              <And>
		                                <Gt><FieldRef Name='ID'></FieldRef><Value Type='Number'>{0}</Value></Gt>
			                            <Lt><FieldRef Name='ID'></FieldRef><Value Type='Number'>{1}</Value></Lt>
		                              </And>
                                        <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq>
                                    </And>
                                </Where>
                            </Query>
                            <ViewFields>
                                <FieldRef Name='Title' />
                            </ViewFields>
                        </View>", lowerId, upperId);
            return qry;
        }

        public static CamlQuery CreateLargeCheckedInFilesQuery(int lowerId, int upperId, int rowLimit)
        {
            // Create the large list query for checked IN files
            // <IsNull><FieldRef Name='CheckoutUser' LookupId='TRUE'></FieldRef></IsNull>
            // between /viewfields and /view <RowLimit>{2}</RowLimit> but took it out because it forces the item position to be null prematurely
            var qry = new CamlQuery();
            qry.ViewXml = string.Format(@"
                        <View Scope='RecursiveAll'>
                            <Query>
                                <Where>
	                              <And>
		                              <And>
		                                <Gt><FieldRef Name='ID'></FieldRef><Value Type='Number'>{0}</Value></Gt>
			                            <Lt><FieldRef Name='ID'></FieldRef><Value Type='Number'>{1}</Value></Lt>
		                              </And>
		                              <And>
                                        <IsNull><FieldRef Name='CheckoutUser' LookupId='TRUE'></FieldRef></IsNull>
			                            <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq>
		                              </And>
                                    </And>
                                </Where>
                            </Query>
                            <ViewFields>
                                <FieldRef Name='Title' />
                            </ViewFields>
                        </View>", lowerId, upperId);
            return qry;
        }

        public static CamlQuery CreateLargeCheckedOutFilesQuery(int lowerId, int upperId, int rowLimit)
        {
            // Create the large list query for checked OUT files
            var qry = new CamlQuery();
            qry.ViewXml = string.Format(@"
                        <View Scope='RecursiveAll'>
                            <Query>
                                <Where>
	                              <And>
		                              <And>
		                                <Gt><FieldRef Name='ID'></FieldRef><Value Type='Number'>{0}</Value></Gt>
			                            <Lt><FieldRef Name='ID'></FieldRef><Value Type='Number'>{1}</Value></Lt>
		                              </And>
		                              <And>
                                        <IsNotNull>
                                            <FieldRef Name='CheckoutUser' LookupId='TRUE'></FieldRef>
                                        </IsNotNull>
			                            <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq>
		                              </And>
                                    </And>
                                </Where>
                            </Query>
                            <ViewFields>
                                <FieldRef Name='Title' />
                            </ViewFields>
                        </View>", lowerId, upperId);
            return qry;
        }

//        public static void InventoryListItems(string siteAddress, string listTitle, int rowLimit, Utils.ActingUser actingUser)
//        {
//            ClientContext ctx = new ClientContext(siteAddress);
//            Web currentWeb = ctx.Web;
//            int lowerId = -1;
//            int upperId = rowLimit;
//            List<string> list = new List<string>();

//            if (string.IsNullOrEmpty(actingUser.UserLoginName))
//            {
//                ctx.Credentials = CredentialCache.DefaultCredentials;
//            }
//            else
//            {
//                if (actingUser.UserLoginName.IndexOf("@") != -1)
//                {
//                    ctx.Credentials = new SharePointOnlineCredentials(actingUser.UserLoginName,
//                        actingUser.UserPassword);
//                }
//                else
//                {
//                    ctx.Credentials = new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword,
//                        actingUser.UserDomain);
//                }
//            }
//            List tmpList = currentWeb.Lists.GetByTitle(listTitle); 
//            ctx.Load(tmpList);
//            ctx.ExecuteQuery();

//            int folderCount = 0;
//            int checkedInCount = 0;
//            int checkedOutCount = 0;
//            if (tmpList.BaseType == BaseType.DocumentLibrary)
//            {
//                // Get a count of folders to be removed from the total list item count for comparing to checked in docs
//                var folders = tmpList.GetItems(CreateAllFoldersQuery());
//                ctx.Load(folders, icol => icol.Include(i => i.File));
//                ctx.ExecuteQuery();
//                folderCount = folders.Count;

//                // Get the checked in files from the list
//                var itemsCheckedIn = tmpList.GetItems(CreateCheckedOutFilesQuery());
//                ctx.Load(itemsCheckedIn, icol => icol.Include(i => i.File, i => i.DisplayName));
//                ctx.ExecuteQuery();
//                checkedInCount = itemsCheckedIn.Count;


//                // Get the checked out files from the list
//                var itemsCheckedOut = tmpList.GetItems(CreateCheckedOutFilesQuery());
//                ctx.Load(itemsCheckedOut, icol => icol.Include(i => i.File, i => i.DisplayName));
//                ctx.ExecuteQuery();
//                checkedOutCount = itemsCheckedOut.Count;
//            }

//            ListItemCollectionPosition itemPosition = null;
//            do
//            {
//                CamlQuery camlQuery = new CamlQuery();
//                camlQuery.ListItemCollectionPosition = itemPosition;

//                string viewXml = string.Format(@"
//                        <View Scope='RecursiveAll'>
//                            <Query>
//                                <Where>
//	                              <And>
//		                              <And>
//		                                <Gt><FieldRef Name='ID'></FieldRef><Value Type='Number'>{0}</Value></Gt>
//			                            <Lt><FieldRef Name='ID'></FieldRef><Value Type='Number'>{1}</Value></Lt>
//		                              </And>
//                                        <Eq>
//                                            <FieldRef Name='FSObjType' />
//                                            <Value Type='Integer'>0</Value>
//                                        </Eq>
//                                    </And>
//                                </Where>
//                            </Query>
//                            <ViewFields>
//                                <FieldRef Name='Title' />
//                            </ViewFields>
//                            <RowLimit>{2}</RowLimit>
//                        </View>", lowerId, upperId, rowLimit);

//                camlQuery.ViewXml = viewXml;
//                ListItemCollection listItems = tmpList.GetItems(camlQuery);
//                ctx.Load(listItems);
//                ctx.ExecuteQuery();
//                itemPosition = listItems.ListItemCollectionPosition;
//                foreach (ListItem listItem in listItems)
//                {
//                    //// need to see if we need to run query to get ID
//                    //ctx.Load(listItem, li => li.Id);
//                    //ctx.ExecuteQuery();
//                    try
//                    {
//                        list.Add(listItem.Id.ToString());
//                    }
//                    catch (Exception ex31Exception)
//                    {
//                        Console.WriteLine(ex31Exception.Message);
//                    }
//                    //// try to get the file of each list item, and see what its checked out status is
//                    //File file = listItem.File;
//                    //ctx.Load(file, f => f.CheckOutType, f => f.Name);
//                    //ctx.ExecuteQuery();
//                    //if (file.CheckOutType.ToString() == "None")
//                    //{
//                    //    largeCheckedInCount++;
//                    //}
//                }
//                lowerId += rowLimit;
//                upperId += rowLimit;
//                Console.WriteLine("start row{0}  end row{1}", lowerId, upperId);
//            }
//            while (itemPosition != null);

//            foreach (var line in list)
//            {
//                string listFilePath = "C:\\Temp\\LargeListIds.csv";
//                // Write the information about large lists to the inventory CSV file
//                string[] passingListObject = new string[2];
//                passingListObject[0] = listFilePath;
//                passingListObject[1] = line;
//                WriteReports.WriteText(passingListObject);
//            }
//            //Utils.SpinAnimation.Stop();
//            //Console.WriteLine();
//            //Console.WriteLine("Could not retrieve any more details for " + listTitle + ", because it is a large list.");
//            //Utils.SpinAnimation.Start();
//        }

//        private async Task<List<ListItem>> GetListItems(string siteAddress, string listTitle, int rowLimit)
//        {
//            List<ListItem> items = new List<ListItem>();
//            ClientContext ctx = new ClientContext(siteAddress);
//            List list = ctx.Web.Lists.GetByTitle(listTitle);
//            //int rowLimit = 100;
//            ListItemCollectionPosition position = null;

//            string viewXml = string.Format(@"
//                        <View Scope='RecursiveAll'>
//                            <Query>
//                                <Where>
//                                    <Eq>
//                                        <FieldRef Name='FSObjType' />
//                                        <Value Type='Integer'>0</Value>
//                                    </Eq>
//                                </Where>
//                            </Query>
//                            <ViewFields>
//                                <FieldRef Name='Title' />
//                            </ViewFields>
//                            <RowLimit>{0}</RowLimit>
//                        </View>", rowLimit);

//            var camlQuery = new CamlQuery();
//            camlQuery.ViewXml = viewXml;

//            do
//            {
//                ListItemCollection listItems = null;
//                if (listItems != null && listItems.ListItemCollectionPosition != null)
//                {
//                    camlQuery.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
//                }

//                listItems = list.GetItems(camlQuery);
//                ctx.Load(listItems);
//                //Task contextTask = ctx.ExecuteQueryAsync();
//                //await Task.WhenAll(contextTask);

//                //Task contextTask = ctx.ExecuteQuery();
//                ctx.ExecuteQuery();
//                await Task.WhenAll();
//                position = listItems.ListItemCollectionPosition;
//                items.AddRange(listItems.ToList());
//            }
//            while (position != null);
//            return items;
//        }

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
    }
}

using System;
using System.Linq;
using System.Net;
using Microsoft.SharePoint.Client;

namespace MNIT.Inventory
{
    public class GetInfoPath
    {
        // Method to inventory information about sites with workflows and instances of workflows
        public static void InventoryInfoPath(string siteAddress, Utilities.ActingUser actingUser, ref int infoPathFormCounter, ref int infoPathExternalConnCounter, string csvFilePath)
        {
            ClientContext ctx = new ClientContext(siteAddress);

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
            Site site = ctx.Site;
            // Load web and web properties
            ctx.Load(subWeb, w => w.Webs, w => w.Url, w => w.Title, w => w.Lists, w => w.Id);
            // Execute Query against web
            ctx.ExecuteQuery();
            try
            {
                string currentWebTitle = subWeb.Title;
                string currentWebUrl = subWeb.Url;
                Uri tempUri = new Uri(currentWebUrl);
                string urlDomain = tempUri.Host;
                string urlProtocol = tempUri.Scheme;
                string siteCollId = "";
                string webId = "";
                // find the SCAs or owners of the site collection
                ctx.Load(site, sc => sc.Owner, sc => sc.RootWeb, sc => sc.Id);
                ctx.ExecuteQuery();
                string rootWebOwner = site.Owner.Email;
                if (string.IsNullOrEmpty(rootWebOwner))
                {
                    rootWebOwner = site.Owner.Title;
                }
                // Only get the web ID and the Site Collection Web ID if it is not an App web
                if (currentWebUrl.ElementAt(8) != 'a')
                {
                    webId = subWeb.Id.ToString();
                    siteCollId = site.Id.ToString();
                }

                foreach (List tmpList in subWeb.Lists)
                {
                    // Variables
                    string infoPathFormName = "";
                    string hasExternalConnections = "";
                    // Load list and list properties
                    ctx.Load(tmpList, t => t.Title, t => t.DefaultViewUrl, t => t.Id, t => t.BaseType, t => t.BaseTemplate, t => t.HasExternalDataSource, t => t.DefaultNewFormUrl, t => t.Hidden, t => t.IsPrivate);
                    // Execute Query against the list
                    ctx.ExecuteQuery();
                    string currentListTitle = "";
                    // Build the URL
                    string currentListUrl = urlProtocol + "://" + urlDomain + tmpList.DefaultViewUrl;
                    // Build the Web Application Name
                    string webApplication = urlDomain.Split('.')[0];
                    // Count and list all Form libraries with InfoPath forms, and add to the form counter for the rollup report
                    if (tmpList.BaseType == BaseType.DocumentLibrary && tmpList.BaseTemplate == 115)
                    {
                        // Add to the counter for the rollup file
                        infoPathFormCounter++;
                        infoPathFormName = "Form Library";
                        currentListTitle = tmpList.Title;
                        if (tmpList.HasExternalDataSource == true)
                        {
                            infoPathExternalConnCounter++;
                            hasExternalConnections = "Yes";
                        }
                    }
                    // Count and list all lists with InfoPath for custom list forms, and add to the form counter for the rollup report
                    if (tmpList.BaseType != BaseType.DocumentLibrary && tmpList.Hidden == false)
                    {
                        // Query the list to find a folder called Item that might contain an item called template.xsn
                        FolderCollection listFolders = tmpList.RootFolder.Folders;
                        ctx.Load(listFolders);
                        ctx.ExecuteQuery();
                        foreach (var listFolder in listFolders)
                        {
                            string customFormFolder = listFolder.Name;
                            if (customFormFolder == "Item")
                            {
                                FileCollection customForms = listFolder.Files;
                                ctx.Load(customForms);
                                ctx.ExecuteQuery();
                                foreach (var customForm in customForms)
                                {
                                    ctx.Load(customForm, cf => cf.Name);
                                    ctx.ExecuteQuery();
                                    if (customForm.Name.Contains("xsn"))
                                    {
                                        infoPathFormCounter++;
                                        infoPathFormName = "Customized List";
                                        currentListTitle = tmpList.Title;
                                        if (tmpList.HasExternalDataSource == true)
                                        {
                                            infoPathExternalConnCounter++;
                                            hasExternalConnections = "Yes";
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if(!string.IsNullOrEmpty(currentListTitle))
                    {
                        // Write the List Data to the detailed CSV file
                        string[] passingListObject = new string[11];
                        passingListObject[0] = csvFilePath;
                        passingListObject[1] = webApplication;
                        passingListObject[2] = siteCollId;
                        passingListObject[3] = webId;
                        passingListObject[4] = currentWebTitle;
                        passingListObject[5] = currentWebUrl;
                        passingListObject[6] = rootWebOwner;
                        passingListObject[7] = currentListTitle;
                        passingListObject[8] = currentListUrl;
                        passingListObject[9] = infoPathFormName;
                        passingListObject[10] = hasExternalConnections;
                        WriteReports.WriteText(passingListObject);
                    }
                }

                // Recursively inventory InfoPath forms for all sub webs; Only use webs and sub webs that are not a host for apps
                foreach (var recursiveSubWeb in ctx.Web.Webs)
                {
                    if (recursiveSubWeb.Url.ElementAt(8) != 'a')
                    {
                        InventoryInfoPath(recursiveSubWeb.Url, actingUser, ref infoPathFormCounter, ref infoPathExternalConnCounter, csvFilePath);
                    }
                }
            }
            catch (Exception ex06Exception)
            {
                Utilities.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Could not access all data for site {0}, {1}. {2}", subWeb.Title, subWeb.Url, ex06Exception.Message);
                Utilities.SpinAnimation.Start();
            }
            finally
            {
                ctx.Dispose();
            }
        }
    }
}

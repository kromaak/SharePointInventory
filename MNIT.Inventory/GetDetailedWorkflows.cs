using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;

namespace MNIT.Inventory
{
    public class GetDetailedWorkflows
    {
        // Method to inventory information about sites with workflows and instances of workflows
        public static void InventoryWorkflowsDetailed(string siteAddress, Utilities.ActingUser actingUser, ref int runningInstancesCounter, string csvFilePath)
        {
            ClientContext ctx = new ClientContext(siteAddress);
            //ctx.Credentials = new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword, actingUser.UserDomain);
            ctx.Credentials = !string.IsNullOrEmpty(actingUser.UserLoginName) ? new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword, actingUser.UserDomain) : CredentialCache.DefaultCredentials;
            Web subWeb = ctx.Web;
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
                string wfPlatform = "SPD2013";
                string siteCollId = null;
                //string rootWebId = null;
                string webId = null;
                // find the SCAs or owners of the site collection
                Site currentSite = ctx.Site;
                ctx.Load(currentSite, sc => sc.Owner, sc => sc.RootWeb, sc => sc.Id);
                ctx.ExecuteQuery();
                string rootWebOwner = currentSite.Owner.Email;
                if (string.IsNullOrEmpty(rootWebOwner))
                {
                    rootWebOwner = currentSite.Owner.Title;
                }
                // Only get the web ID and the Site Collection Web ID if it is not an App web
                if (currentWebUrl.ElementAt(8) != 'a')
                {
                    //rootWebId = currentSite.RootWeb.Id.ToString();
                    webId = subWeb.Id.ToString();
                    siteCollId = currentSite.Id.ToString();
                }

                // Build out the calls to get information about WorkFlows
                // Workflow Services Manager which will handle all the workflow interaction.
                WorkflowServicesManager wfManager = new WorkflowServicesManager(ctx, subWeb);
                // connect to the instance service
                WorkflowInstanceService wfInstanceService = wfManager.GetWorkflowInstanceService();

                foreach (List tmpList in subWeb.Lists)
                {
                    // Initialize variables
                    int running = 0;
                    string strRunningCount = null;
                    //string strRunningInstances = null;
                    // Load list and list properties
                    ctx.Load(tmpList, t => t.Title, t => t.DefaultViewUrl, t => t.Id, t => t.WorkflowAssociations);
                    // Execute Query against the list
                    ctx.ExecuteQuery();
                    Guid currentListId = tmpList.Id;
                    string currentListTitle = tmpList.Title;
                    string currentListUrl = urlProtocol + "://" + urlDomain + tmpList.DefaultViewUrl;

                    // Connect to WF subscription service
                    WorkflowSubscriptionService wfSubscriptionService = wfManager.GetWorkflowSubscriptionService();
                    // Create WF Subscription Collection
                    WorkflowSubscriptionCollection wfSubscriptions = wfSubscriptionService.EnumerateSubscriptionsByList(tmpList.Id);
                    // Load the WF subscriptions collection
                    ctx.Load(wfSubscriptions);
                    // Execute the query
                    ctx.ExecuteQuery();
                    // Collect WF information about 2013 WFs, including running instances

                    if (wfSubscriptions.Count > 0)
                    {
                        // Initialize platform or WF version variable
                        wfPlatform = "SPD2013";
                        foreach (var wfSubscription in wfSubscriptions)
                        {
                            // Variables
                            Guid instWfSubscriptionId = new Guid();
                            // Load information about the WF subscription
                            ctx.Load(wfSubscription, wfSub => wfSub.Name, wfSub => wfSub.Id);
                            // Execute the query
                            ctx.ExecuteQuery();
                            Guid wfSubscriptionId = wfSubscription.Id;
                            string wfSubscriptionName = wfSubscription.Name;

                            ListItemCollection items = tmpList.GetItems(CamlQuery.CreateAllItemsQuery());
                            // Load list items
                            ctx.Load(items);
                            // Execute Querty against list items
                            ctx.ExecuteQuery();
                            foreach (ListItem listItem in items)
                            {
                                // Initialize Variables
                                running = 0;
                                // Load list items
                                ctx.Load(listItem, i => i.Id);
                                // Execute the query to retrieve list items
                                ctx.ExecuteQuery();

                                // Enumerate all the instances for each list item
                                var workflowInstances = wfInstanceService.EnumerateInstancesForListItem(currentListId,
                                    listItem.Id);
                                // Load list item WF instance collection
                                ctx.Load(workflowInstances);
                                // Execute Query against WF instances
                                ctx.ExecuteQuery();

                                foreach (var instance in workflowInstances)
                                {
                                    // Load list item WF instance
                                    ctx.Load(instance, n => n.Status, n => n.WorkflowSubscriptionId);
                                    // Execute Query against list item WF instance
                                    ctx.ExecuteQuery();
                                    instWfSubscriptionId = instance.WorkflowSubscriptionId;
                                    // Get the state of the workflow, whether completed, terminated, running, etc
                                    if (wfSubscriptionId == instWfSubscriptionId)
                                    {
                                        // if there is an instance of the WF in a running state add to the running counter
                                        if (instance.Status.ToString().ToLower() == "started" ||
                                            instance.Status.ToString().ToLower() == "running")
                                        {
                                            running++;
                                            runningInstancesCounter++;
                                        }
                                    }
                                }
                            }
                            strRunningCount = running.ToString();
                            // Write the 2013 WF information about the site, the list, the workflow association, and the workflow instance to the inventory CSV file
                            //WriteToStream(siteCollId, webId, currentWebTitle, currentWebUrl, rootWebOwner, currentListTitle, currentListUrl,
                            //    wfPlatform, wfSubscriptionName, strRunningCount, null, streamWriter);
                            string[] passingWfObject = new string[11];
                            passingWfObject[0] = csvFilePath;
                            passingWfObject[1] = siteCollId;
                            passingWfObject[2] = webId;
                            passingWfObject[3] = currentWebTitle;
                            passingWfObject[4] = currentWebUrl;
                            passingWfObject[5] = rootWebOwner;
                            passingWfObject[6] = currentListTitle;
                            passingWfObject[7] = currentListUrl;
                            passingWfObject[8] = wfPlatform;
                            passingWfObject[9] = wfSubscriptionName;
                            passingWfObject[10] = strRunningCount;
                            WriteReports.WriteText(passingWfObject);
                        }
                    }

                    // Collect WF information about 2010 WFs
                    foreach (var association in tmpList.WorkflowAssociations)
                    {
                        string wfAssocName = "";
                        // Load WF associations and WF association properties
                        ctx.Load(association, a => a.Name, a => a.InstantiationUrl, a => a.Id, a => a.BaseId);
                        // Execute Query against workflow associations
                        ctx.ExecuteQuery();
                        wfAssocName = association.Name;
                        wfPlatform = "SPD2010";
                        string associationUrl = "";
                        if (!string.IsNullOrEmpty(association.InstantiationUrl))
                        {
                            associationUrl = association.InstantiationUrl.ToLower();
                            if (associationUrl.Contains("nintexworkflow"))
                            {
                                wfPlatform = "NINTEX";
                            }
                        }
                        strRunningCount = "Not avail for 2010 WFs";

                        // Do not document each previous WF association, only capture information about the currently published WF association
                        if (!wfAssocName.Contains("Previous Version"))
                        {
                            // Write 2010 WF the information about the site, the list, the workflow association, and the workflow instance to the inventory CSV file
                            //WriteToStream(siteCollId, webId, currentWebTitle, currentWebUrl, rootWebOwner, currentListTitle, currentListUrl, wfPlatform, wfAssocName, strRunningCount, null, streamWriter);
                            string[] passingWfObject = new string[11];
                            passingWfObject[0] = csvFilePath;
                            passingWfObject[1] = siteCollId;
                            passingWfObject[2] = webId;
                            passingWfObject[3] = currentWebTitle;
                            passingWfObject[4] = currentWebUrl;
                            passingWfObject[5] = rootWebOwner;
                            passingWfObject[6] = currentListTitle;
                            passingWfObject[7] = currentListUrl;
                            passingWfObject[8] = wfPlatform;
                            passingWfObject[9] = wfAssocName;
                            passingWfObject[10] = strRunningCount;
                            WriteReports.WriteText(passingWfObject);
                        }
                    }
                }
            }
            catch (Exception ex02Exception)
            {
                Utilities.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Could not access all data for site {0}, {1}. {2}", subWeb.Title, subWeb.Url, ex02Exception.Message);
                Utilities.SpinAnimation.Start();
            }

            // Recursively inventory workflow instances for sub webs; Only use webs and sub web that are not a host for apps
            foreach (var recursiveSubWeb in ctx.Web.Webs)
            {
                if (recursiveSubWeb.Url.ElementAt(8) != 'a')
                {
                    InventoryWorkflowsDetailed(recursiveSubWeb.Url, actingUser, ref runningInstancesCounter, csvFilePath);
                }
            }
        }
    }
}

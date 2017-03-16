using System;
using System.Linq;
using System.Net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;

namespace MNIT.Inventory
{
    public class GetStandardWorkflows
    {        // Method to inventory information about sites with workflows and instances of workflows
        public static void InventoryWorkflowsStandard(string siteAddress, Utilities.ActingUser actingUser, ref int nintexCounter, ref int spd2010Counter, ref int spd2013Counter, string csvFilePath)
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
            Site siteCollection = ctx.Site;
            // Load web and web properties
            ctx.Load(ctx.Web, w => w.Webs, w => w.Url, w => w.Title, w => w.Lists, w => w.Id, w => w.ContentTypes);
            // Execute Query against web
            ctx.ExecuteQuery();
            try
            {
                string currentWebTitle = subWeb.Title;
                string currentWebUrl = subWeb.Url;
                Uri tempUri = new Uri(currentWebUrl);
                string urlDomain = tempUri.Host;
                string urlProtocol = tempUri.Scheme;
                string wfPlatform = "";
                string siteCollId = "";
                string webId = "";
                // find the SCAs or owners of the site collection
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
                    webId = subWeb.Id.ToString();
                    siteCollId = siteCollection.Id.ToString();
                }
                // Build the Web Application Name
                string webApplication = urlDomain.Split('.')[0];
                
                // Build out the calls to get information about WorkFlows
                // Workflow Services Manager which will handle all the workflow interaction.
                WorkflowServicesManager wfManager = new WorkflowServicesManager(ctx, subWeb);
                // Listing out workflows associated with Lists and Libraries
                foreach (List tmpList in subWeb.Lists)
                {
                    // Load list and list properties
                    ctx.Load(tmpList, t => t.Title, t => t.DefaultViewUrl, t => t.Id, t => t.WorkflowAssociations, t => t.TemplateFeatureId);
                    // Execute Query against the list
                    ctx.ExecuteQuery();
                    string currentListTitle = tmpList.Title;
                    // Build the URL
                    string currentListUrl = urlProtocol + "://" + urlDomain + tmpList.DefaultViewUrl;
                    string associationUrl = "";
                    // Connect to WF subscription service
                    WorkflowSubscriptionService wfSubscriptionService = wfManager.GetWorkflowSubscriptionService();
                    // Create WF Subscription Collection
                    WorkflowSubscriptionCollection wfSubscriptions = wfSubscriptionService.EnumerateSubscriptionsByList(tmpList.Id);
                    // Load the WF subscriptions collection
                    ctx.Load(wfSubscriptions);
                    // Execute the query
                    ctx.ExecuteQuery();
                    // Collect WF information about 2013 WFs
                    foreach (var wfSubscription in wfSubscriptions)
                    {
                        ctx.Load(wfSubscription, wfSub => wfSub.Name, wfSub => wfSub.Id);
                        ctx.ExecuteQuery();
                        string wfSubscriptionName = wfSubscription.Name;
                        string wfSubscriptionId = wfSubscription.Id.ToString();
                        wfPlatform = "SPD2013";
                        if (!wfSubscriptionName.Contains("Previous Version"))
                        {
                            spd2013Counter++;
                        }
                        //// Write the 2013 WF information about the site, the list, the workflow association, and the workflow instance to the inventory CSV file
                        //WriteToStream(siteCollId, webId, currentWebTitle, currentWebUrl, rootWebOwner, currentListTitle, currentListUrl,
                        //    wfPlatform, wfSubscriptionName, wfSubscriptionId, null, streamWriter);
                        string[] passingWfObject = new string[12];
                        passingWfObject[0] = csvFilePath;
                        passingWfObject[1] = webApplication;
                        passingWfObject[2] = siteCollId;
                        passingWfObject[3] = webId;
                        passingWfObject[4] = currentWebTitle;
                        passingWfObject[5] = currentWebUrl;
                        passingWfObject[6] = rootWebOwner;
                        passingWfObject[7] = currentListTitle;
                        passingWfObject[8] = currentListUrl;
                        passingWfObject[9] = wfPlatform;
                        passingWfObject[10] = wfSubscriptionName;
                        passingWfObject[11] = wfSubscriptionId;
                        WriteReports.WriteText(passingWfObject);
                    }

                    // Collect WF information about 2010 WFs
                    foreach (var association in tmpList.WorkflowAssociations)
                    {
                        // Initialize Variables
                        string wfAssocName = "";
                        //string wfAssocType = "";
                        // Load WF associations and WF association properties
                        ctx.Load(association, a => a.Name, a => a.Id);
                        // Execute Query against workflow associations
                        ctx.ExecuteQuery();
                        wfAssocName = association.Name;
                        wfPlatform = "SPD2010";
                        string associationId = association.Id.ToString();
                        if (!string.IsNullOrEmpty(association.InstantiationUrl))
                        {
                            associationUrl = association.InstantiationUrl.ToLower();
                            if (associationUrl.Contains("nintexworkflow"))
                            {
                                wfPlatform = "NINTEX";
                                if (!wfAssocName.Contains("Previous Version"))
                                {
                                    nintexCounter++;
                                }
                            }
                            else
                            {
                                if (!wfAssocName.Contains("Previous Version"))
                                {
                                    spd2010Counter++;
                                }
                            }
                        }

                        // Do not document each previous WF association, only capture information about the currently published WF association
                        if (!wfAssocName.Contains("Previous Version"))
                        {
                            //// Write 2010 WF the information about the site, the list, the workflow association, and the workflow instance to the inventory CSV file
                            //WriteToStream(siteCollId, webId, currentWebTitle, currentWebUrl, rootWebOwner, currentListTitle, currentListUrl, wfPlatform, wfAssocName, associationId, null, streamWriter);
                            string[] passingWfObject = new string[12];
                            passingWfObject[0] = csvFilePath;
                            passingWfObject[1] = webApplication;
                            passingWfObject[2] = siteCollId;
                            passingWfObject[3] = webId;
                            passingWfObject[4] = currentWebTitle;
                            passingWfObject[5] = currentWebUrl;
                            passingWfObject[6] = rootWebOwner;
                            passingWfObject[7] = currentListTitle;
                            passingWfObject[8] = currentListUrl;
                            passingWfObject[9] = wfPlatform;
                            passingWfObject[10] = wfAssocName;
                            passingWfObject[11] = associationId;
                            WriteReports.WriteText(passingWfObject);
                        }
                    }
                }
                
                // Listing out workflows associated with Content Types
                foreach (ContentType tmpCtype in subWeb.ContentTypes)
                {
                    // Load Content Types and Content Type properties
                    ctx.Load(tmpCtype, ct => ct.Name, ct => ct.WorkflowAssociations, ct => ct.Id);
                    // Execute Query against the list
                    ctx.ExecuteQuery();
                    string currentCtypeName = tmpCtype.Name;
                    // Build the URL
                    //string currentCtypeUrl = urlProtocol + "://" + urlDomain + tmpCtype.DefaultViewUrl;
                    string associationUrl = "";
                    //// Build the Web Application Name
                    //string webApplication = urlDomain.Split('.')[0];
                    // Connect to WF subscription service
                    WorkflowSubscriptionService wfSubscriptionService = wfManager.GetWorkflowSubscriptionService();
                    // Create WF Subscription Collection
                    WorkflowSubscriptionCollection wfSubscriptions = wfSubscriptionService.EnumerateSubscriptions();
                        //wfSubscriptionService.EnumerateSubscriptionsByList(tmpCtype.Id);
                    // Load the WF subscriptions collection
                    ctx.Load(wfSubscriptions);
                    // Execute the query
                    ctx.ExecuteQuery();
                    // Collect WF information about 2013 WFs
                    foreach (var wfSubscription in wfSubscriptions)
                    {
                        ctx.Load(wfSubscription, wfSub => wfSub.Name, wfSub => wfSub.Id);
                        ctx.ExecuteQuery();
                        string wfSubscriptionName = wfSubscription.Name;
                        string wfSubscriptionId = wfSubscription.Id.ToString();
                        wfPlatform = "SPD2013";
                        if (!wfSubscriptionName.Contains("Previous Version"))
                        {
                            spd2013Counter++;
                        }
                        //// Write the 2013 WF information about the site, the Content Type, the workflow association, and the workflow instance to the inventory CSV file
                        string[] passingWfObject = new string[12];
                        passingWfObject[0] = csvFilePath;
                        passingWfObject[1] = webApplication;
                        passingWfObject[2] = siteCollId;
                        passingWfObject[3] = webId;
                        passingWfObject[4] = currentWebTitle;
                        passingWfObject[5] = currentWebUrl;
                        passingWfObject[6] = rootWebOwner;
                        passingWfObject[7] = currentCtypeName;
                        passingWfObject[8] = "Content Type Workflow";
                        passingWfObject[9] = wfPlatform;
                        passingWfObject[10] = wfSubscriptionName;
                        passingWfObject[11] = wfSubscriptionId;
                        WriteReports.WriteText(passingWfObject);
                    }

                    // Collect WF information about 2010 WFs
                    foreach (var association in tmpCtype.WorkflowAssociations)
                    {
                        // Initialize Variables
                        string wfAssocName = "";
                        //string wfAssocType = "";
                        // Load WF associations and WF association properties
                        ctx.Load(association, a => a.Name, a => a.Id);
                        // Execute Query against workflow associations
                        ctx.ExecuteQuery();
                        wfAssocName = association.Name;
                        wfPlatform = "SPD2010";
                        string associationId = association.Id.ToString();
                        if (!string.IsNullOrEmpty(association.InstantiationUrl))
                        {
                            associationUrl = association.InstantiationUrl.ToLower();
                            if (associationUrl.Contains("nintexworkflow"))
                            {
                                wfPlatform = "NINTEX";
                                if (!wfAssocName.Contains("Previous Version"))
                                {
                                    nintexCounter++;
                                }
                            }
                            else
                            {
                                if (!wfAssocName.Contains("Previous Version"))
                                {
                                    spd2010Counter++;
                                }
                            }
                        }

                        // Do not document each previous WF association, only capture information about the currently published WF association
                        if (!wfAssocName.Contains("Previous Version"))
                        {
                            // Write 2010 WF the information about the site, the list, the workflow association, and the workflow instance to the inventory CSV file
                            string[] passingWfObject = new string[12];
                            passingWfObject[0] = csvFilePath;
                            passingWfObject[1] = webApplication;
                            passingWfObject[2] = siteCollId;
                            passingWfObject[3] = webId;
                            passingWfObject[4] = currentWebTitle;
                            passingWfObject[5] = currentWebUrl;
                            passingWfObject[6] = rootWebOwner;
                            passingWfObject[7] = currentCtypeName;
                            passingWfObject[8] = "Content Type Workflow";
                            passingWfObject[9] = wfPlatform;
                            passingWfObject[10] = wfAssocName;
                            passingWfObject[11] = associationId;
                            WriteReports.WriteText(passingWfObject);
                        }
                    }
                }
                if (siteCollection.RootWeb.Url == currentWebUrl)
                {
                    // connect to the deployment service
                    var workflowDeploymentService = wfManager.GetWorkflowDeploymentService();
                    // get all installed workflows
                    var publishedWorkflowDefinitions = workflowDeploymentService.EnumerateDefinitions(true);
                    ctx.Load(publishedWorkflowDefinitions);
                    ctx.ExecuteQuery();
                    foreach (var definition in publishedWorkflowDefinitions)
                    {
                        //Console.WriteLine(definition.DisplayName);
                        ctx.Load(definition, wfDef => wfDef.DisplayName, wfDef => wfDef.Id, wfDef => wfDef.AssociationUrl, wfDef => wfDef.Published);
                        ctx.ExecuteQuery();
                        // Do not document each previous WF association, only capture information about the currently published WF association
                        //if (!wfAssocName.Contains("Previous Version"))
                        //{
                            // Write 2010 WF the information about the site, the list, the workflow association, and the workflow instance to the inventory CSV file
                            string[] passingWfObject = new string[12];
                            passingWfObject[0] = csvFilePath;
                            passingWfObject[1] = webApplication;
                            passingWfObject[2] = siteCollId;
                            passingWfObject[3] = webId;
                            passingWfObject[4] = currentWebTitle;
                            passingWfObject[5] = currentWebUrl;
                            passingWfObject[6] = rootWebOwner;
                            passingWfObject[7] = definition.DisplayName;
                            passingWfObject[8] = "Content Type Workflow";
                            passingWfObject[9] = "SPD2010";
                            passingWfObject[10] = definition.AssociationUrl;
                            passingWfObject[11] = definition.Id.ToString();
                            WriteReports.WriteText(passingWfObject);
                        //}
                    }
                }
            }
            catch (Exception ex03Exception)
            {
                Utilities.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Could not access all data for site {0}, {1}. {2}", subWeb.Title, subWeb.Url, ex03Exception.Message);
                Utilities.SpinAnimation.Start();
            }

            // Recursively inventory standard workflow info for all sub webs; Only use webs and sub webs that are not a host for apps
            foreach (var recursiveSubWeb in ctx.Web.Webs)
            {
                if (recursiveSubWeb.Url.ElementAt(8) != 'a')
                {
                    InventoryWorkflowsStandard(recursiveSubWeb.Url, actingUser, ref nintexCounter, ref spd2010Counter, ref spd2013Counter, csvFilePath);
                }
            }
        }
    }
}

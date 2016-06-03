﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;

using Utils = MNIT.Utilities;

namespace MNIT.Inventory
{
    public class GetWebs
    {
        // Method to inventory webs, web owners, templates, and sandbox solutions
        public static void InventoryWebs(string siteAddress, string exportedWp, Utilities.ActingUser actingUser, ref int siteTemplateCounter, ref int solutionCounter, ref int masterPageCounter, ref int pageLayoutCounter, ref int customPageCounter, ref int appCounter, ref int dropoffCounter, ref int listTemplateCounter, ref int exportedWpCounter, string csvFilePath)
        {
            // Variables
            string strSolutionCount = "";
            string strListOfListTemplates = "";
            string sandboxGalleryUrl = "";
            string listGalleryUrl = "";
            string rootWebOwner = "";
            string urlTemplate = "";
            string spVersion = "";
            string spSiteType = "SP Web";
            string customSiteMaster = "";
            string customSystemMaster = "";
            string siteCollId = "";
            string webId = "";
            string rootFolder = "";
            string pageLayouts = "";
            string strSiteSize = "";
            string strWebCount = "";
            //string requestAccess = "";
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
            Web subWeb = ctx.Web;
            // Load web and web properties
            ctx.Load(subWeb, w => w.Webs, w => w.Url, w => w.Title, w => w.Lists, w => w.WebTemplate, w => w.Id, w => w.MasterUrl, w => w.CustomMasterUrl, w => w.ServerRelativeUrl);
            // Execute Query against web
            ctx.ExecuteQuery();

            try
            {
                // Get the site template used by each web in the site collection
                urlTemplate = subWeb.WebTemplate;
                // Find the SCAs or owners of the site collection
                Site siteCollection = ctx.Site;
                ctx.Load(siteCollection, sc => sc.Owner, sc => sc.Url, sc => sc.RootWeb, sc => sc.RequiredDesignerVersion, sc => sc.CompatibilityLevel, sc => sc.Id, sc => sc.Usage);
                ctx.ExecuteQuery();
                // Callout the primary owner of the site collection
                rootWebOwner = siteCollection.Owner.Email;
                if (string.IsNullOrEmpty(rootWebOwner))
                {
                    rootWebOwner = siteCollection.Owner.Title;
                }
                // site collection size
                long siteSize = siteCollection.Usage.Storage;
                Int64 conversionSize = Convert.ToInt64(siteSize);
                string currentWebUrl = subWeb.Url;
                webId = subWeb.Id.ToString();
                Uri tempUri = new Uri(currentWebUrl);
                string urlDomain = tempUri.Host;
                // Build the Web Application Name
                string webApplication = urlDomain.Split('.')[0];
                siteCollId = siteCollection.Id.ToString();
                spVersion = siteCollection.CompatibilityLevel.ToString();
                // Find all the webs in the site collection that have a custom master page applied
                string masterPage = null;
                string customPage = null;
                int index1 = subWeb.MasterUrl.LastIndexOf('/');
                if (index1 != -1)
                {
                    masterPage = subWeb.MasterUrl.Substring(index1);
                }
                int index2 = subWeb.CustomMasterUrl.LastIndexOf('/');
                if (index2 != -1)
                {
                    customPage = subWeb.CustomMasterUrl.Substring(index2);
                }
                if (!string.IsNullOrEmpty(masterPage) && !string.IsNullOrEmpty(customPage))
                {
                    if ((masterPage != "/seattle.master" && masterPage != "/oslo.master") || (customPage != "/seattle.master" && customPage != "/oslo.master"))
                    {
                        //customMaster = masterPage.Replace("/", "");
                        masterPageCounter++;
                    }
                    if (masterPage != "/seattle.master" && masterPage != "/oslo.master")
                    {
                        customSiteMaster = masterPage.Replace("/", "");
                    }
                    if (customPage != "/seattle.master" && customPage != "/oslo.master")
                    {
                        customSystemMaster = customPage.Replace("/", "");
                    }
                }

                //// Need to use something like this to streamline the query process
                //// Need to replace or update foreach and if(templateId == 850) mechanisms 
                //ListCollection listCollection = subWeb.Lists;
                //ctx.Load(listCollection, lc => lc.Include(list => list.Title, list => list.DefaultViewUrl, list => list.ItemCount, list => list.BaseTemplate, list => list.BaseType).Where(list => list.BaseTemplate == 850));
                
                // Find all publishing pages libraries with pages that have custom layouts assigned
                foreach (List tmpList in subWeb.Lists)
                {
                    // Load list and list properties
                    ctx.Load(tmpList, t => t.Title, t => t.DefaultViewUrl, t => t.ItemCount, t => t.BaseTemplate, t => t.BaseType);
                    // Execute Query against the list
                    ctx.ExecuteQuery();
                    // Publishing Pages library has no value for ListTemplate, but has a ListTemplateID of 850
                    int templateId = tmpList.BaseTemplate;
                    //string templateBaseType = tmpList.BaseType.ToString();
                    if (templateId == 850)
                    {

                        // Create a new file path for adding custom page layouts to a new custom report to act against
                        string customPagesPath = csvFilePath.Replace("Webs", "Pages");
                        
                        string[] argsStrings = new string[4];
                        argsStrings[0] = siteAddress;
                        argsStrings[1] = customPagesPath;
                        argsStrings[2] = tmpList.Title;
                        argsStrings[3] = exportedWp;
                        GetPages.InventoryPages(argsStrings, ref pageLayoutCounter, ref customPageCounter, ref exportedWpCounter);
                    }
                }

                //// Track Access Request settings
                //if (!string.IsNullOrEmpty(subWeb.RequestAccessEmail))
                //{
                //    requestAccess = subWeb.RequestAccessEmail;
                //}
                // Count all the MPS templates to find Meeting Workspace Sites
                //if (subWeb.WebTemplate == "MPS" || subWeb.WebTemplate == "MPS" || subWeb.WebTemplate == "MPS" || subWeb.WebTemplate == "MPS" || subWeb.WebTemplate == "MPS" || subWeb.WebTemplate == "MPS" || subWeb.WebTemplate == "MPS")
                if (subWeb.WebTemplate == "MPS")
                {
                    siteTemplateCounter++;
                }
                if (String.Equals(subWeb.Url, siteCollection.RootWeb.Url, StringComparison.CurrentCultureIgnoreCase))
                {
                    sandboxGalleryUrl = siteCollection.RootWeb.Url + "/_catalogs/solutions/";
                    List solutionGallery = subWeb.GetList(sandboxGalleryUrl);
                    ctx.Load(solutionGallery, sg => sg.ItemCount);
                    ctx.ExecuteQuery();
                    solutionCounter = solutionGallery.ItemCount;
                    strSolutionCount = solutionCounter.ToString();
                    spSiteType = "SP Root Web";
                    // get a count of list templates in the list template gallery
                    listGalleryUrl = siteCollection.RootWeb.Url + "/_catalogs/lt/";
                    List listTemplateGallery = subWeb.GetList(listGalleryUrl);
                    ctx.Load(listTemplateGallery, lg => lg.ItemCount);
                    ctx.ExecuteQuery();
                    listTemplateCounter = listTemplateGallery.ItemCount;
                    if (listTemplateCounter > 0)
                    {
                        // Compile a list of List Templates in the List Gallery to report on
                        ListItemCollection items = listTemplateGallery.GetItems(CamlQuery.CreateAllItemsQuery());
                        // Load list items
                        ctx.Load(items);
                        // Execute Querty against list items
                        ctx.ExecuteQuery();
                        foreach (var listItem in items)
                        {
                            ctx.Load(listItem, lti => lti.DisplayName);
                            ctx.ExecuteQuery();
                            strListOfListTemplates += "; " + listItem.DisplayName;
                        }
                        //strListOfListTemplates = Enumerable.Aggregate(items, strListOfListTemplates, (current, listItem) => current + ("; " + listItem.DisplayName));
                        listGalleryUrl += strListOfListTemplates;
                    }
                    // Get the basic info about the root web like sub web count
                    int subWebCount = subWeb.Webs.Count;
                    strWebCount = subWebCount.ToString();
                    //strSiteSize = siteSize.ToString();
                    strSiteSize = SizeSuffix(conversionSize);
                }
                // Find all the webs that have a root folder called DropOffLibrary
                string webUrl = subWeb.ServerRelativeUrl;
                string routingRulesUrl = webUrl.Equals("/", StringComparison.Ordinal)
                    ? "/RoutingRules"
                    : webUrl + "/RoutingRules";
                List routingRules = null;
                ListCollection lists = subWeb.Lists;
                IQueryable<List> queryObjects = lists.Include(list => list.RootFolder).Where(list => list.RootFolder.ServerRelativeUrl == routingRulesUrl);
                IEnumerable<List> filteredLists = ctx.LoadQuery(queryObjects);
                ctx.ExecuteQuery();
                routingRules = filteredLists.FirstOrDefault();
                if (routingRules != null)
                {
                    dropoffCounter++;
                    rootFolder = "Drop Off Library";
                }
                // Write a line for each web
                string[] passingWebObject = new string[19];
                passingWebObject[0] = csvFilePath;
                passingWebObject[1] = webApplication;
                passingWebObject[2] = siteCollId;
                passingWebObject[3] = webId;
                passingWebObject[4] = subWeb.Title;
                passingWebObject[5] = subWeb.Url;
                passingWebObject[6] = rootWebOwner;
                passingWebObject[7] = urlTemplate;
                passingWebObject[8] = sandboxGalleryUrl;
                passingWebObject[9] = strSolutionCount;
                passingWebObject[10] = spVersion;
                passingWebObject[11] = spSiteType;
                passingWebObject[12] = customSiteMaster;
                passingWebObject[13] = customSystemMaster;
                passingWebObject[14] = rootFolder;
                passingWebObject[15] = pageLayouts;
                passingWebObject[16] = listGalleryUrl;
                passingWebObject[17] = strSiteSize;
                passingWebObject[18] = strWebCount;
                //passingWebObject[16] = accessRequest;
                WriteReports.WriteText(passingWebObject);

                // For every web, look for sub webs and Only use webs that are not a host for apps
                foreach (var recursiveSubWeb in ctx.Web.Webs)
                {
                    if (recursiveSubWeb.Url.ElementAt(8) != 'a')
                    {
                        InventoryWebs(recursiveSubWeb.Url, exportedWp, actingUser, ref siteTemplateCounter, ref solutionCounter, ref masterPageCounter, ref pageLayoutCounter, ref customPageCounter, ref appCounter, ref dropoffCounter, ref listTemplateCounter, ref exportedWpCounter, csvFilePath);
                    }
                    // Try to inventory the app webs that would otherwise be missed in the inventory
                    else
                    {
                        appCounter++;
                        // Write App Web Object To CSV
                        string[] passingAppWebObject = new string[11];
                        passingAppWebObject[0] = csvFilePath;
                        passingAppWebObject[1] = siteCollId;
                        passingAppWebObject[2] = webId;
                        passingAppWebObject[3] = recursiveSubWeb.Title;
                        passingAppWebObject[4] = recursiveSubWeb.Url;
                        passingAppWebObject[5] = "";
                        passingAppWebObject[6] = urlTemplate;
                        passingAppWebObject[7] = "";
                        passingAppWebObject[8] = "";
                        passingAppWebObject[9] = "15";
                        passingAppWebObject[10] = "APP WEB";
                        WriteReports.WriteText(passingAppWebObject);
                    }
                }
            }
            catch (Exception ex01Exception)
            {
                // Add a line in for each web for reference
                //Console.WriteLine(@"Could not access all data for site {0}, {1}. {2}", subWeb.Title, subWeb.Url, ex01Exception.Message);
                Console.WriteLine(@"Could not access all data for site {0}, {1}. {2}", subWeb.Title, subWeb.Url, ex01Exception);
            }
            finally
            {
                ctx.Dispose();
            }
        }

        public static readonly string[] SizeSuffixes = { "bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB" };
        public static string SizeSuffix(Int64 value)
        {
            if (value < 0) { return "-" + SizeSuffix(-value); }
            if (value == 0) { return "0.0 bytes"; }

            int mag = (int)Math.Log(value, 1024);
            decimal adjustedSize = (decimal)value / (1L << (mag * 10));

            return string.Format("{0:n1} {1}", adjustedSize, SizeSuffixes[mag]);
        }

    }
}
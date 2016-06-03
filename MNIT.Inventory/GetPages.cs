using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;

namespace MNIT.Inventory
{
    class GetPages
    {
        public static void InventoryPages(string[] args, ref int pageLayoutCounter, ref int customPageCounter, ref int exportedWpCounter)
        {
            string siteAddress = args[0];
            //string customPagesPath = args[1];
            ClientContext ctx = new ClientContext(siteAddress);
            Web subWeb = ctx.Web;
            Site siteCollection = ctx.Site;
            string siteCollId = "";
            string webId = "";

            // Load web and web properties
            ctx.Load(subWeb, w => w.Url, w => w.Title, w => w.Id);
            // Execute Query against web
            ctx.ExecuteQuery();
            // find the SCAs or owners of the site collection
            ctx.Load(siteCollection, sc => sc.Owner, sc => sc.Id);
            // Execute Query against site collection
            ctx.ExecuteQuery();

            webId = subWeb.Id.ToString();
            siteCollId = siteCollection.Id.ToString();
            string currentWebUrl = subWeb.Url;
            Uri tempUri = new Uri(currentWebUrl);
            string urlDomain = tempUri.Host;
            // Build the Web Application Name
            string webApplication = urlDomain.Split('.')[0];
            ListCollection listCollection = subWeb.Lists;
            List tmpList = listCollection.GetByTitle(args[2]);



            string pageLayouts = "";
            string exportedWp = "";
            if (!string.IsNullOrEmpty(args[3]))
            {
                exportedWp = args[3];
            }
            // Create a new file path for adding custom page layouts to a new custom report to act against
            string customPagesPath = args[1].Replace("Webs", "Pages");
            // Get all items from the Pages library
            ListItemCollection items = tmpList.GetItems(CamlQuery.CreateAllItemsQuery());
            // Load list items
            ctx.Load(items);
            // Execute Querty against list items
            ctx.ExecuteQuery();
            foreach (ListItem listItem in items)
            {
                // Variables
                string listItemName = "";
                string pageUrl = "";
                string pageLayoutUrl = "";
                string pageLayoutDesc = "";
                string webPartTitle = "";
                string webPartTitleUrl = "";
                //string pageAuthorName = "";
                //string pageAuthorEmail = "";
                //string pageEditorName = "";
                //string pageEditorEmail = "";
                //string pageModifier = "";
                ctx.Load(listItem, p => p.DisplayName);
                ctx.ExecuteQuery();
                // Get the Page Layout from each page
                FieldUrlValue pageLayoutLink = (listItem["PublishingPageLayout"]) as FieldUrlValue;

                // get current page file properties
                File file = listItem.File;
                //ctx.Load(file, cp => cp.CustomizedPageStatus, cp => cp.Author, cp => cp.ModifiedBy);
                ctx.Load(file, cp => cp.CustomizedPageStatus);
                ctx.ExecuteQuery();

                if (pageLayoutLink != null)
                {
                    // Get the file extension from each object to make sure it is a web page
                    string pageExtension = (listItem["FileLeafRef"]) as String;
                    if (!string.IsNullOrEmpty(pageExtension))
                    {
                        // Build the string that represents the full URL of the page
                        pageUrl = siteAddress + "/" + args[2] + "/" + pageExtension;
                    }
                    pageLayoutUrl = pageLayoutLink.Url;
                    pageLayoutDesc = pageLayoutLink.Description;
                    //string pageLayoutInfo = "";
                    if (pageLayoutUrl.ToLower().IndexOf("advancedsearchlayout") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("advancedsearchresults") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("articlelinks") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("articleright") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("blankwebpartpage") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("defaultlayout") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("enterprisewiki") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("errorlayout") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("pagefromdoclayout") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("searchmain") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("searchresults") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("welcomelinks") > -1 ||
                        pageLayoutUrl.ToLower().IndexOf("welcometoc") > -1)
                    {
                        // Do nothing - we expected these normal OOB page layouts
                    }
                    else
                    {
                        // Report on customized pages that have been detached from a page layout
                        if (pageLayoutUrl.ToLower().IndexOf("disconnectedpublishingpage") > -1)
                        {
                            pageLayoutUrl = "DETACHED FROM LAYOUT";
                            customPageCounter++;
                        }
                        else
                        {
                            // If the page layout is something non-standard, report on it
                            if (!string.IsNullOrEmpty(pageLayoutDesc))
                            {
                                // Add the page layout Description to the report if it is not null 
                                pageLayouts += "[" + pageLayoutDesc + "] ";
                            }
                            pageLayoutCounter++;
                        }
                        listItemName = listItem.DisplayName;
                    }
                    #region pagemodifier properties
                    //// Get the file editor for each object
                    //if (!string.IsNullOrEmpty(file.Author.LoginName))
                    //{
                    //    if (Utils.SpObjects.UserExists(siteAddress, file.Author.LoginName, actingUser) ==
                    //        true)
                    //    {
                    //        User pageAuthUser = file.Author;
                    //        pageAuthorName = pageAuthUser.Title;
                    //        pageAuthorEmail = pageAuthUser.Email;
                    //    }
                    //}
                    //if (!string.IsNullOrEmpty(file.ModifiedBy.LoginName))
                    //{
                    //    if (
                    //        Utils.SpObjects.UserExists(siteAddress, file.ModifiedBy.LoginName, actingUser) ==
                    //        true)
                    //    {
                    //        User pageEditUser = file.ModifiedBy;
                    //        pageEditorName = pageEditUser.Title;
                    //        pageEditorEmail = pageEditUser.Email;
                    //    }
                    //}
                    //if (!string.IsNullOrEmpty(pageAuthorName))
                    //{
                    //    pageModifier += "Created By: " + pageAuthorName + "; ";
                    //}
                    //if (!string.IsNullOrEmpty(pageAuthorEmail))
                    //{
                    //    pageModifier += pageAuthorEmail + "; ";
                    //}
                    //if (!string.IsNullOrEmpty(pageEditorName))
                    //{
                    //    pageModifier += "Modified By: " + pageEditorName + "; ";
                    //}
                    //if (!string.IsNullOrEmpty(pageEditorEmail))
                    //{
                    //    pageModifier += pageEditorEmail + ";";
                    //}
                    #endregion
                    // get specific web parts from current page
                    if (!string.IsNullOrEmpty(exportedWp))
                    {
                        LimitedWebPartManager lwpmShared =
                            file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                        LimitedWebPartManager lwpmUser =
                            file.GetLimitedWebPartManager(PersonalizationScope.User);

                        WebPartDefinitionCollection webPartDefinitionCollectionShared = lwpmShared.WebParts;
                        WebPartDefinitionCollection webPartDefinitionCollectionUser = lwpmUser.WebParts;

                        ctx.Load(webPartDefinitionCollectionShared,
                            w => w.Include(wp => wp.WebPart, wp => wp.Id));
                        ctx.Load(webPartDefinitionCollectionUser,
                            w => w.Include(wp => wp.WebPart, wp => wp.Id));

                        ctx.Load(subWeb, p => p.Url);
                        ctx.ExecuteQuery();

                        foreach (WebPartDefinition webPartDefinition in webPartDefinitionCollectionShared)
                        {
                            WebPart webPart = webPartDefinition.WebPart;
                            ctx.Load(webPart, wp => wp.ZoneIndex, wp => wp.Properties, wp => wp.Title,
                                wp => wp.Subtitle, wp => wp.TitleUrl);
                            ctx.ExecuteQuery();

                            // Check to see if the webpart title url matches the site url
                            // Need to continue to build this out, but it may not give perfect results in the end
                            // because a user might configure any url as the webpart title url, not necessarily the link to an object in the current site 
                            if (webPart.Title.IndexOf(exportedWp) > -1)
                            {
                                exportedWpCounter++;
                                listItemName = listItem.DisplayName;
                                webPartTitle = webPart.Title;
                                webPartTitleUrl = webPart.TitleUrl;

                                if (!string.IsNullOrEmpty(webPartTitleUrl))
                                {
                                    string listFullUrl = "";
                                    if (webPartTitleUrl.StartsWith("/"))
                                    {
                                        Uri wpTempUri = new Uri(subWeb.Url);
                                        string wpUrlDomain = wpTempUri.Host;
                                        string wpUrlProtocol = wpTempUri.Scheme;
                                        listFullUrl = wpUrlProtocol + "://" + wpUrlDomain + webPartTitleUrl;
                                    }
                                    if (webPartTitleUrl.StartsWith("http"))
                                    {
                                        listFullUrl = webPartTitleUrl;
                                    }

                                    //List exportedWebPartList = subWeb.GetList(listFullUrl);

                                    ListCollection listCollectionExpWebPart = subWeb.Lists;
                                    // This will not work if the link is pointing at a non-default view of a list
                                    // Need to add function to check for, and trim off anything after a /forms/ piece of text in the link
                                    //ctx.Load(listCollectionExpWebPart, lewp => lewp.Include(list => list.DefaultViewUrl).Where(list => list.DefaultViewUrl.Contains(webPartTitleUrl)));
                                    //ctx.ExecuteQuery();
                                    ctx.Load(listCollectionExpWebPart, all => all
                                      .Where(l => l.RootFolder.Name == listFullUrl)
                                      .Include(l => l.Id));
                                    ctx.ExecuteQuery();
                                    //List expList = listCollectionExpWebPart.Single();

                                    if (listCollectionExpWebPart.Count > 0)
                                    {
                                        foreach (var expList in listCollectionExpWebPart)
                                        {
                                            ctx.Load(expList, el => el.Id);
                                            ctx.ExecuteQuery();
                                            Utilities.SpinAnimation.Stop();
                                            Console.WriteLine("{0} List exists... {1}", webPartTitle, expList.Id);
                                            Utilities.SpinAnimation.Start();
                                        }
                                    }
                                    else
                                    {
                                        Utilities.SpinAnimation.Stop();
                                        Console.WriteLine("List {0} does not exist in current site on page {1}", listFullUrl, pageUrl);
                                        Utilities.SpinAnimation.Start();
                                    }
                                    //Console.WriteLine(pageUrl + " WP title url: " + webPartTitleUrl);
                                    //Console.WriteLine(subWeb.ServerRelativeUrl);
                                }
                                //else
                                //{
                                //    Console.WriteLine("No, could not get URL to match{0}, or wp title url was empty: {1}", pageUrl, webPartTitleUrl);
                                //}
                            }
                        }
                    }
                }
                else
                {
                    if (!Utilities.SpObjects.ObjectIsFolder(listItem))
                    {
                        // Add non-SharePoint pages to the report
                        // Get the file extension from each object to make sure it is a web page
                        string pageExtension = (listItem["FileLeafRef"]) as String;
                        if (!string.IsNullOrEmpty(pageExtension))
                        {
                            // Build the string that represents the full URL of the page
                            pageUrl = siteAddress + "/" + args[2] + "/" + pageExtension;
                        }
                        string fileType = pageExtension.Substring(pageExtension.LastIndexOf('.') + 1);
                        if (fileType.ToLower() == "aspx")
                        {
                            pageLayoutDesc = fileType + " file, or a Wiki Page";
                        }
                        else
                        {
                            pageLayoutDesc = fileType + " file, not a SP page";
                        }
                        pageLayoutUrl = "None";
                        customPageCounter++;
                        listItemName = listItem.DisplayName;
                    }

                }

                // Write the line to the pages report
                if (!string.IsNullOrEmpty(listItemName))
                {
                    string[] passingDetailedPagesObject = new string[9];
                    passingDetailedPagesObject[0] = customPagesPath;
                    passingDetailedPagesObject[1] = webApplication;
                    passingDetailedPagesObject[2] = siteCollId;
                    passingDetailedPagesObject[3] = webId;
                    passingDetailedPagesObject[4] = pageUrl;
                    passingDetailedPagesObject[5] = webPartTitle;
                    passingDetailedPagesObject[6] = pageLayoutDesc;
                    passingDetailedPagesObject[7] = pageLayoutUrl;
                    passingDetailedPagesObject[8] = webPartTitleUrl;
                    //passingDetailedPagesObject[5] = pageModifier;
                    WriteReports.WriteText(passingDetailedPagesObject);
                }
            }
        }
    }
}

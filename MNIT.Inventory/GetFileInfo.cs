using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;

namespace MNIT.Inventory
{
    class GetFileInfo
    {
    }

    public class GetCheckedOutFiles
    {
        private Boolean ObjectIsCheckedOut(string siteAddress, string fileName)
        {
            bool checkedOutDoc = true;
            //string requestAccess = "";
            ClientContext ctx = new ClientContext(siteAddress);
            //ctx.Credentials = !string.IsNullOrEmpty(actingUser.UserLoginName) ? new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword, actingUser.UserDomain) : System.Net.CredentialCache.DefaultCredentials;
            Web subWeb = ctx.Web;
            // Load web and web properties
            ctx.Load(subWeb, w => w.Lists);
            // Execute Query against web
            ctx.ExecuteQuery();
            foreach (List tmpList in subWeb.Lists)
            {
                // Initialize variables
                // Load list and list properties
                ctx.Load(tmpList, t => t.Title, t => t.DefaultViewUrl, t => t.ItemCount, t => t.EnableVersioning,
                    t => t.EnableMinorVersions, t => t.IsPrivate, t => t.Hidden, t => t.MajorVersionLimit,
                    t => t.MajorWithMinorVersionsLimit, t => t.BaseType, t => t.ForceCheckout);
                // Execute Query against the list
                ctx.ExecuteQuery();

                ListItemCollection items = tmpList.GetItems(CamlQuery.CreateAllItemsQuery());
                // Load list items
                ctx.Load(items);
                // Execute Querty against list items
                ctx.ExecuteQuery();
                foreach (ListItem listItem in items)
                {
                    ctx.Load(listItem, co => co.DisplayName);
                    ctx.ExecuteQuery();
                    // get current page file properties
                    File file = listItem.File;
                    //ctx.Load(file, f => f.Author, f => f.ModifiedBy);
                    ctx.Load(file, f => f.CheckedOutByUser, f => f.CheckOutType);
                    ctx.ExecuteQuery();

                    if (!string.IsNullOrEmpty(file.CheckedOutByUser.LoginName))
                    {
                        checkedOutDoc = true;
                    }
                    else
                    {
                        checkedOutDoc = false;
                    }
                    // add items to checked out documents report
                }
            }

            return checkedOutDoc;
        }
        
    }
    public class GetWebPartsOnPage
    {
        //private static Boolean ObjectIsFolder(string siteAddress, string fileName)
        //{

        //    ClientContext ctx = new ClientContext(siteAddress);
        //    ctx.Credentials = !string.IsNullOrEmpty(actingUser.UserLoginName) ? new NetworkCredential(actingUser.UserLoginName, actingUser.UserPassword, actingUser.UserDomain) : System.Net.CredentialCache.DefaultCredentials;
        //    Web subWeb = ctx.Web;
        //    // get specific web parts from current page
        //    File file = listItem.File;
        //    ctx.Load(file);
        //    ctx.ExecuteQuery();
        //    LimitedWebPartManager lwpmShared = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
        //    LimitedWebPartManager lwpmUser = file.GetLimitedWebPartManager(PersonalizationScope.User);

        //    WebPartDefinitionCollection webPartDefinitionCollectionShared = lwpmShared.WebParts;
        //    WebPartDefinitionCollection webPartDefinitionCollectionUser = lwpmUser.WebParts;

        //    ctx.Load(webPartDefinitionCollectionShared, w => w.Include(wp => wp.WebPart, wp => wp.Id));
        //    ctx.Load(webPartDefinitionCollectionUser, w => w.Include(wp => wp.WebPart, wp => wp.Id));

        //    ctx.Load(subWeb, p => p.Url);
        //    ctx.ExecuteQuery();

        //    foreach (WebPartDefinition webPartDefinition in webPartDefinitionCollectionShared)
        //    {
        //        WebPart webPart = webPartDefinition.WebPart;
        //        ctx.Load(webPart, wp => wp.ZoneIndex, wp => wp.Properties, wp => wp.Title, wp => wp.Subtitle, wp => wp.TitleUrl);
        //        ctx.ExecuteQuery();

        //        ////Once the webPart is loaded, you can do your modification as follows
        //        //webPart.Title = "My New Web Part Title";
        //        //webPartDefinition.SaveWebPartChanges();
        //        //ctx.ExecuteQuery();
        //        if (webPart.Title.IndexOf("MN.IT Finder") > -1)
        //        {
        //            exportedWpCounter++;
        //            listItemName = listItem.DisplayName;
        //            webPartTitle = webPart.Title;
        //        }
        //    }
        //}
        
    }
}

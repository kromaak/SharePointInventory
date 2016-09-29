using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Net;
using Utils = MNIT.Utilities;

namespace MNIT.Inventory
{
    internal class Program
    {
        //private static void Main(string[] args)
        //{
        //}
    }

    public class GetUsersAndGroups
    {
        // Method to inventory information about sites with workflows and instances of workflows
        public static void InventoryAdGroups(string siteAddress, Utils.ActingUser actingUser, ref string adGroups, ref string permissionLevels, string action, string csvFilePath)
        {
            // ClientContext declaration
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
            Site siteCollection = ctx.Site;
            // Load web and web properties
            ctx.Load(subWeb, w => w.Webs, w => w.Url, w => w.Title, w => w.Lists, w => w.Id, w => w.RoleDefinitions);
            // Execute Query against web
            ctx.ExecuteQuery();
            // find the SCAs or owners of the site collection
            ctx.Load(siteCollection, sc => sc.Owner, sc => sc.RootWeb, sc => sc.Id);
            ctx.ExecuteQuery();
            string currentWebUrl = subWeb.Url;
            Uri tempUri = new Uri(currentWebUrl);
            string urlDomain = tempUri.Host;
            // Build the Web Application Name
            string webApplication = urlDomain.Split('.')[0];
            string siteCollId = siteCollection.Id.ToString();
            if (siteCollection.RootWeb.Url == subWeb.Url)
            {
                RoleDefinitionCollection roleDefinitionColl = subWeb.RoleDefinitions;
                foreach (RoleDefinition roleDefinition in roleDefinitionColl)
                {
                    if (roleDefinition.Name != "Full Control" && roleDefinition.Name != "Design" &&
                        roleDefinition.Name != "Edit" && roleDefinition.Name != "Contribute" &&
                        roleDefinition.Name != "Read" && roleDefinition.Name != "Limited Access" &&
                        roleDefinition.Name != "View Only" && roleDefinition.Name != "Approve" &&
                        roleDefinition.Name != "Manage Hierarchy" && roleDefinition.Name != "Restricted Read" &&
                        roleDefinition.Name != "Restricted Interfaces for Translation")
                    {
                        if (!permissionLevels.Contains(roleDefinition.Name))
                        {
                            permissionLevels += roleDefinition.Name + "; ";
                        }
                    }
                }
            }
            try
            {
                // variables for user properties
                //string userDistinguishedName = "";
                //string groupType = "";
                string userEmail = "";
                string userTitle = "";
                string userLoginName = "";
                // Get all AD groups not in SPGroups for web
                UserCollection adUserGroupColl = subWeb.SiteUsers;
                ctx.Load(adUserGroupColl);
                ctx.ExecuteQuery();
                // Create a file to store Partner user info if Partners exist
                string partnerCsvFilePath = csvFilePath.Replace("User", "PartnerUser");

                foreach (User adGrp in adUserGroupColl)
                {
                    ctx.Load(adGrp, adg => adg.PrincipalType, adg => adg.Id, adg => adg.UserId);
                    ctx.ExecuteQuery();

                    if (action.ToLower() == "groups")
                    {
                        if (!adGroups.Contains(adGrp.Title) && adGrp.PrincipalType == PrincipalType.SecurityGroup)
                        {
                            if (adGroups.Length > 31000)
                            {
                                // Write AD Object To CSV early to avoid the 32759 char limit in Excel cells
                                string[] passingAdObjects = new string[6];
                                passingAdObjects[0] = csvFilePath;
                                passingAdObjects[1] = webApplication;
                                passingAdObjects[2] = siteCollId;
                                //passingAdObjects[3] = subWeb.Url;
                                passingAdObjects[3] = currentWebUrl;
                                passingAdObjects[4] = adGroups;
                                passingAdObjects[5] = permissionLevels;
                                WriteReports.WriteText(passingAdObjects);
                                // Empty the adGroups reference, and start the count over
                                adGroups = String.Empty;
                            }
                            userEmail = !string.IsNullOrEmpty(adGrp.Email) ? "[" + adGrp.Email + "]" : "";
                            userTitle = "[" + adGrp.Title + "]";
                            userLoginName = "[" + adGrp.LoginName + "]";
                            //userLoginName = "[" + adGrp.LoginName + "][" + adGrp.UserId + "][" + adGrp.Id + "]";
                            adGroups += userTitle + userLoginName + userEmail + "; ";
                        }
                    }
                    if (action.ToLower() == "users")
                    {
                        if (!adGroups.Contains(adGrp.Title) && adGrp.PrincipalType == PrincipalType.User)
                        {
                            if (adGroups.Length > 31000)
                            {
                                // Write AD Object To CSV early to avoid the 32759 char limit in Excel cells
                                string[] passingAdObjects = new string[6];
                                passingAdObjects[0] = csvFilePath;
                                passingAdObjects[1] = webApplication;
                                passingAdObjects[2] = siteCollId;
                                //passingAdObjects[3] = subWeb.Url;
                                passingAdObjects[3] = currentWebUrl;
                                passingAdObjects[4] = adGroups;
                                passingAdObjects[5] = permissionLevels;
                                WriteReports.WriteText(passingAdObjects);
                                // Empty the adGroups reference, and start the count over
                                adGroups = String.Empty;
                            }
                            userEmail = !string.IsNullOrEmpty(adGrp.Email) ? "[" + adGrp.Email + "]" : "";
                            userTitle = "[" + adGrp.Title + "]";
                            userLoginName = "[" + adGrp.LoginName + "]";
                            adGroups += userTitle + userLoginName + userEmail + "; ";
                            // Additionaly Add Partner Users to a separate CSV on a per site collection per line
                            if (userLoginName.IndexOf("partner\\") > -1)
                            {
                                // Write Partner AD Object To CSV
                                string[] passingPartnerAdObjects = new string[7];
                                passingPartnerAdObjects[0] = partnerCsvFilePath;
                                passingPartnerAdObjects[1] = webApplication;
                                passingPartnerAdObjects[2] = siteCollId;
                                //passingPartnerAdObjects[3] = subWeb.Url;
                                passingPartnerAdObjects[3] = currentWebUrl;
                                passingPartnerAdObjects[4] = adGrp.Title;
                                passingPartnerAdObjects[5] = adGrp.LoginName;
                                passingPartnerAdObjects[6] = adGrp.Email;
                                WriteReports.WriteText(passingPartnerAdObjects);
                            }
                        }
                    }
                }

                // Get all SP Groups for web
                GroupCollection collGroup = subWeb.SiteGroups;
                ctx.Load(collGroup, gc => gc.Include(grp => grp.Users));
                ctx.ExecuteQuery();

                foreach (Group siteGroup in collGroup)
                {
                    UserCollection collUser = siteGroup.Users;
                    foreach (User siteUser in collUser)
                    {
                        ctx.Load(siteUser, su => su.PrincipalType);
                        ctx.ExecuteQuery();
                        if (action.ToLower() == "groups")
                        {
                            if (!adGroups.Contains(siteUser.Title) && siteUser.PrincipalType == PrincipalType.SecurityGroup)
                            {
                                if (adGroups.Length > 31000)
                                {
                                    // Write AD Object To CSV early to avoid the 32759 char limit in Excel cells
                                    string[] passingAdObjects = new string[6];
                                    passingAdObjects[0] = csvFilePath;
                                    passingAdObjects[1] = webApplication;
                                    passingAdObjects[2] = siteCollId;
                                    //passingAdObjects[3] = subWeb.Url;
                                    passingAdObjects[3] = currentWebUrl;
                                    passingAdObjects[4] = adGroups;
                                    passingAdObjects[5] = permissionLevels;
                                    WriteReports.WriteText(passingAdObjects);
                                    // Empty the adGroups reference, and start the count over
                                    adGroups = String.Empty;
                                }
                                userEmail = !string.IsNullOrEmpty(siteUser.Email) ? "[" + siteUser.Email + "]" : "";
                                userTitle = "[" + siteUser.Title + "]";
                                userLoginName = "[" + siteUser.LoginName + "]";
                                //userLoginName = "[" + siteUser.LoginName + "][" + siteUser.UserId + "][" + siteUser.Id + "]";
                                adGroups += userTitle + userLoginName + userEmail + "; ";
                            }
                        }
                        if (action.ToLower() == "users")
                        {
                            if (!adGroups.Contains(siteUser.Title) && siteUser.PrincipalType == PrincipalType.User)
                            {
                                if (adGroups.Length > 31000)
                                {
                                    // Write AD Object To CSV early to avoid the 32759 char limit in Excel cells
                                    string[] passingAdObjects = new string[6];
                                    passingAdObjects[0] = csvFilePath;
                                    passingAdObjects[1] = webApplication;
                                    passingAdObjects[2] = siteCollId;
                                    //passingAdObjects[3] = subWeb.Url;
                                    passingAdObjects[3] = currentWebUrl;
                                    passingAdObjects[4] = adGroups;
                                    passingAdObjects[5] = permissionLevels;
                                    WriteReports.WriteText(passingAdObjects);
                                    // Empty the adGroups reference, and start the count over
                                    adGroups = String.Empty;
                                }
                                userEmail = !string.IsNullOrEmpty(siteUser.Email) ? "[" + siteUser.Email + "]" : "";
                                userTitle = "[" + siteUser.Title + "]";
                                userLoginName = "[" + siteUser.LoginName + "]";
                                adGroups += userTitle + userLoginName + userEmail + "; ";
                                // Additionaly Add Partner Users to a separate CSV on a per site collection per line
                                if (userLoginName.IndexOf("partner\\") > -1)
                                {
                                    // Write Partner AD Object To CSV
                                    string[] passingPartnerAdObjects = new string[7];
                                    passingPartnerAdObjects[0] = partnerCsvFilePath;
                                    passingPartnerAdObjects[1] = webApplication;
                                    passingPartnerAdObjects[2] = siteCollId;
                                    //passingPartnerAdObjects[3] = subWeb.Url;
                                    passingPartnerAdObjects[3] = currentWebUrl;
                                    passingPartnerAdObjects[4] = siteUser.Title;
                                    passingPartnerAdObjects[5] = siteUser.LoginName;
                                    passingPartnerAdObjects[6] = siteUser.Email;
                                    WriteReports.WriteText(passingPartnerAdObjects);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex05Exception)
            {
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Could not access all data for site {0}, {1}.{2}",
                    subWeb.Title,
                    subWeb.Url,
                    ex05Exception.Message);
                Utils.SpinAnimation.Start();
            }
        }
    }
}

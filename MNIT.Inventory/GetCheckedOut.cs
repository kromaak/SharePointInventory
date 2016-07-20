using System;
using Microsoft.SharePoint.Client;

using Utils = MNIT.Utilities;

namespace MNIT.Inventory
{
    class GetCheckedOut
    {
        public static void InventoryCheckedOutDocs(string[] args, ref string strFileCount, ref string strFolderCount, ref string strCheckedInCount, ref string strCheckedOutCount, ref string strNeverCheckedInCount, ref string manageListUrl, ref int largeListCounter, ref int unlimitedVerCounter, ref int siteCollCheckedOut)
        {
            string siteAddress = args[0];
            //string customPagesPath = args[1];
            ClientContext ctx = new ClientContext(siteAddress);
            Web subWeb = ctx.Web;
            ListCollection listCollection = subWeb.Lists;
            string currentListTitle = args[1];
            List tmpList = listCollection.GetByTitle(currentListTitle);
            // Get item count and compare to checked in document count
            // counter including folders
            int totalListItemCount = tmpList.ItemCount;
            int folderCount = 0;
            // counter excluding folders
            int listItemCount = 0;
            int checkedInCount = 0;
            int checkedOutCount = 0;
            int neverCheckedInCount = 0;
            string strTotalListItemCount = "";
            strFileCount = "";
            strFolderCount = "";
            strCheckedInCount = "";
            //strCheckedOutCount = "";
            //strNeverCheckedInCount = "";
            //manageListUrl = "";
            int largeListDiv = 3;
            if (tmpList.BaseType == BaseType.DocumentLibrary)
            {

                // Break the list down into chunks smaller than 5000 to inventory checked out docs
                if (totalListItemCount > 4999)
                {
                    Utilities.SpinAnimation.Stop();
                    Console.WriteLine();
                    Console.WriteLine(@"Checking large list {0}; Count: {1}; Divisor: {2}", tmpList.Title,
                        totalListItemCount, largeListDiv);
                    Utilities.SpinAnimation.Start();

                    DivideLargeList(totalListItemCount, largeListDiv);
                }

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
                    // list title
                    //currentListTitle = tmpList.Title;
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
        }





        public static CamlQuery CreateAllFilesQuery()
        {
            var qry = new CamlQuery();
            qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";
            //qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query><RowLimit></RowLimit></View>";
            return qry;
        }

        public static void DivideLargeList(int itemCount, int divisor)
        {
            //int remainder = 0;
            if (itemCount % divisor >= 5000)
            {
                divisor++;
                DivideLargeList(itemCount, divisor);
            }
            else
            {
                int dividedCount = itemCount / divisor;
                Utils.SpinAnimation.Stop();
                Console.WriteLine();
                Console.WriteLine(@"Large List divided count: {0}; Divisor: {1}; Remainder: {2}", itemCount, divisor, dividedCount);
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

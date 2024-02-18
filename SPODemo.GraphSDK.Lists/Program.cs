using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.IdentityModel.Protocols;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;

namespace SPODemo.GraphSDK.Lists
{
    internal class Program
    {
        async static Task Main(string[] args)
        {
            await Run();

            Console.ReadLine();
        }

        async private static Task Run()
        {
            // The client credentials flow requires that you request the
            // /.default scope, and pre-configure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { ".default" };

            //get client credentials from configuration
            var clientId = ConfigurationManager.AppSettings.Get("clientId");
            var clientSecret = ConfigurationManager.AppSettings.Get("clientSecret");
            var tenantId = ConfigurationManager.AppSettings.Get("tenantId");

            //to find your sideId value: go to https://<your-tenant>.sharepoint.com/sites/<sitecoll>/_api/site and search for <d:Id> property value
            var siteId = "a91c43fd-1802-49be-b833-1c549e24fe2e";
            //to find your listId value: go to your SharePoint list page, go to List Settings page and copy the id in the browser address bar (see the URL template below)
            // https://<your-tenant>.sharepoint.com/sites/<sitecoll>/_layouts/15/listedit.aspx?List=%7B<listId>%7D
            var listId = "C87EC976-3AA4-4CA1-AACE-80D370505163";

            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            ClientSecretCredential clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

            GraphServiceClient graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            // Code snippets are only available for the latest version. Current version is 5.x
            // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
            ListCollectionResponse allLists = await graphClient.Sites[siteId].Lists.GetAsync();

            Console.WriteLine("\n\n All Lists Available");
            foreach (List list in allLists.Value)
            {
                Console.WriteLine(list.Name);
            }

            // Code snippets are only available for the latest version. Current version is 5.x
            // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
            List listMetadda = await graphClient.Sites[siteId].Lists[listId].GetAsync();

            Console.WriteLine("\n\n List Metadata");
            Console.WriteLine(listMetadda.Name);

            // Code snippets are only available for the latest version. Current version is 5.x
            // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
            ListItemCollectionResponse items = await graphClient.Sites[siteId].Lists[listId].Items.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "fields" }; //it will expand with *all* Fields and make those available in item.Fields.AdditionalData dictionary
            });

            Console.WriteLine($"\n\n Response List Items {items.Value.Count}"); // Items.GetAsync will retrieve at most the first 200 items in the list
            foreach (ListItem item in items.Value)
            {
                Console.WriteLine(ItemValueToString(item));
            }

            //get all items in list using Items.GetAsync with PageIterator
            ListItemCollectionResponse allItemsResponse = await graphClient.Sites[siteId].Lists[listId].Items.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "fields($select=Category,Area,Region)" }; //explicit selection of fields to expand
            });

            List<ListItem> allItems = new List<ListItem>();
            var pageIterator = PageIterator<ListItem, ListItemCollectionResponse>.CreatePageIterator(graphClient, allItemsResponse, item =>
            {
                allItems.Add(item);
                return true;
            });

            await pageIterator.IterateAsync(); //trigger page iteration

            Console.WriteLine($"\n\n All List Items {allItems.Count}"); //allItems contains all items in the list (it works only if the list contains up to 5000 items)
            foreach (ListItem item in allItems)
            {
                Console.WriteLine(ItemValueToString(item));
            }
        }

        private static string ItemValueToString(ListItem item)
        {
            object category = string.Empty;
            object area = string.Empty;
            object region = string.Empty;

            item.Fields.AdditionalData.TryGetValue("Category", out category);
            item.Fields.AdditionalData.TryGetValue("Area", out area);
            item.Fields.AdditionalData.TryGetValue("Region", out region);

            return $"{item.Id}: {category} - {area} - {region}");
        }
    }
}

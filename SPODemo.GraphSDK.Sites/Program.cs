using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Configuration;
using System.Threading.Tasks;

namespace SPODemo.GraphSDK.Sites
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

            // Values from app registration
            var clientId = ConfigurationManager.AppSettings.Get("clientId");
            var clientSecret = ConfigurationManager.AppSettings.Get("clientSecret");
            var tenantId = ConfigurationManager.AppSettings.Get("tenantId");

            //to find your sideId value: go to https://<your-tenant>.sharepoint.com/sites/<sitecoll>/_api/site and search for <d:Id> property value
            var siteId = "a91c43fd-1802-49be-b833-1c549e24fe2e";

            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            Site site = await graphClient.Sites[siteId].GetAsync();

            Console.WriteLine(site.Name);
            Console.WriteLine(site.CreatedDateTime);
        }
    }
}

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace Flexinets.Sharepoint
{
    public class DocumentRepository
    {
        private readonly String _clientId;
        private readonly String _clientSecret;
        private readonly String _resource;
        private readonly String _tokenUrl;

        public DocumentRepository(String tenantId, String resourceId, String clientId, String clientSecret)
        {
            _clientId = $"{clientId}@{tenantId}";
            _clientSecret = clientSecret;
            _resource = $"{resourceId}/flexinets.sharepoint.com@{tenantId}";
            _tokenUrl = $"https://accounts.accesscontrol.windows.net/{tenantId}/tokens/OAuth/2";
        }


        /// <summary>
        /// Authenticate and get an access token
        /// </summary>
        /// <returns></returns>
        private async Task<dynamic> GetToken()
        {
            // todo cache the token until it expires maybe...
            var client = new HttpClient();
            var content = new FormUrlEncodedContent(new List<KeyValuePair<String, String>>
            {
                new KeyValuePair<string, string>("grant_type","client_credentials"),
                new KeyValuePair<string, string>("client_id", _clientId ),
                new KeyValuePair<string, string>("client_secret", _clientSecret),
                new KeyValuePair<string, string>("resource", _resource)
            });

            var response = await client.PostAsync(_tokenUrl, content);
            response.EnsureSuccessStatusCode();

            return JObject.Parse(await response.Content.ReadAsStringAsync());
        }


        /// <summary>
        /// Create an authenticated http client with access token
        /// </summary>
        /// <returns></returns>
        private async Task<HttpClient> GetAuthenticatedClient()
        {
            var token = await GetToken();
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {token.access_token}");
            client.DefaultRequestHeaders.Add("Accept", "application/json");
            return client;
        }


        /// <summary>
        /// Get a file stream by file name
        /// </summary>
        /// <param name="fileref"></param>
        /// <returns></returns>
        public async Task<Stream> DownloadDocument(String fileref)
        {
            var url = $"https://flexinets.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('/PortalDocuments')/Files('{fileref}')/$value";  // todo does this need escaping or does sharepoint handle that?
            var client = await GetAuthenticatedClient();
            return await client.GetStreamAsync(url);
        }


        /// <summary>
        /// Get recent and pinned documents
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<DocumentModel>> GetDocuments()
        {
            var documents = await GetRecentDocuments();
            var pinnedDocuments = await GetPinnedDocuments();

            var list = documents.Union(pinnedDocuments, new DocumentModelComparer()).OrderByDescending(o => o.Created);
            return list;
        }


        private async Task<IEnumerable<DocumentModel>> GetRecentDocuments()
        {
            var client = await GetAuthenticatedClient();
            var documentsUrl = "https://flexinets.sharepoint.com/_api/web/lists/GetByTitle('PortalDocuments')/items?$orderby=Created desc&$select=FileLeafRef,FileRef,ID,Created,Pinned,Category&$top=10";

            var documents = new List<DocumentModel>();

            dynamic response = JObject.Parse(await client.GetStringAsync(documentsUrl));
            foreach (var item in response.value)
            {
                documents.Add(new DocumentModel
                {
                    Category = (String)item.Category,
                    Filename = (String)item.FileLeafRef,
                    Path = (String)item.FileRef,
                    Created = (DateTime)item.Created,
                    Id = (Int32)item.ID
                });
            }

            return documents;
        }

        // todo refactor
        private async Task<IEnumerable<DocumentModel>> GetPinnedDocuments()
        {
            var client = await GetAuthenticatedClient();
            var pinnedDocumentsUrl = "https://flexinets.sharepoint.com/_api/web/lists/GetByTitle('PortalDocuments')/items?$orderby=Created desc&$select=ID,FileLeafRef,FileRef,Created,Pinned,Category&$filter=Pinned eq 1";

            var documents = new List<DocumentModel>();

            dynamic response = JObject.Parse(await client.GetStringAsync(pinnedDocumentsUrl));
            foreach (var item in response.value)
            {
                documents.Add(new DocumentModel
                {
                    Category = (String)item.Category,
                    Filename = (String)item.FileLeafRef,
                    Path = (String)item.FileRef,
                    Created = (DateTime)item.Created,
                    Id = (Int32)item.ID
                });
            }

            return documents;
        }


        private class DocumentModelComparer : IEqualityComparer<DocumentModel>
        {
            public bool Equals(DocumentModel x, DocumentModel y)
            {
                return x.Id == y.Id;
            }

            public int GetHashCode(DocumentModel obj)
            {
                return obj.Id;
            }
        }
    }
}



// https://flexinets.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('/PortalDocuments')/Files('Coverage Summary_January 2018.pdf')/$value
// https://flexinets.sharepoint.com/_api/web/lists/GetByTitle('PortalDocuments')/items?$select=Title,Created,Category,Pinned&$orderby=Created desc
// var url = "https://flexinets.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('/PortalDocuments')/Files";
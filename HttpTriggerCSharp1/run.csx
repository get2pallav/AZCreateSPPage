#r "Newtonsoft.Json"
using System; 
using System.Net; 
using OfficeDevPnP.Core;  
using Microsoft.SharePoint.Client;  
using Microsoft.Online.SharePoint.TenantAdministration;  
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client.Publishing;
using Newtonsoft.Json;  

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
   // string baseUrl = "https://tenant.sharepoint.com";
   // string tenantUrl_CONFIG_KEY = "tenantUrl";
    //string managedPath = "/sites/";
    string clientId_CONFIG_KEY = "AppId";
    string clientSecret_CONFIG_KEY = "AppSecret";
    string clientId = System.Environment.GetEnvironmentVariable(clientId_CONFIG_KEY, EnvironmentVariableTarget.Process);
    string clientSecret = System.Environment.GetEnvironmentVariable(clientSecret_CONFIG_KEY, EnvironmentVariableTarget.Process);
    string SPUserName = System.Environment.GetEnvironmentVariable("SPUserName",EnvironmentVariableTarget.Process);
    string SPPassword = System.Environment.GetEnvironmentVariable("SPPassword",EnvironmentVariableTarget.Process);
 //   string tenantUrl = System.Environment.GetEnvironmentVariable(tenantUrl_CONFIG_KEY, EnvironmentVariableTarget.Process);

 console.log("")
    if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
    {
        req.CreateResponse(HttpStatusCode.BadRequest,
           new { body = "Application configuration missing for Function App" });
    }
    string completeUrl = "https://codesignedintranet.sharepoint.com/sites/dev/Communities/SPFx/";
    OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();
    ClientContext siteContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(completeUrl, SPUserName, SPPassword);
    //Tenant tenant = new Tenant(tenantContext);
    try
    {
        var web = siteContext.Web;
        siteContext.Load(web);
        siteContext.ExecuteQuery();
      // Input Parameters  
        string pageName = "TestPage";  
        string pageTemplate = "IntranetNews";  
        string pageTitle = "Test Page";  
        bool toBePublished = false;  
        DateTime startTime = DateTime.Now;  
        DateTime endTime = DateTime.Now;  
        bool scheduled = false;  
/*
        var pageList = web.Lists.GetByTitle("Pages");
        var folder  = web.GetFolderByServerRelativeUrl("/Departments/Development/Pages/News");
        siteContext.Load(pageList);
        siteContext.Load(folder);
        siteContext.ExecuteQuery();
        
        log.Info(folder.Name.ToString());
    */    

        // Adds Publishing Page 

       siteContext.Web.AddPublishingPage( pageName, pageTemplate, pageTitle, toBePublished, null,startTime,endTime,scheduled);  
       // log.Info(pubPage.ToString());
      //  siteContext.Load(pubPage);
      //  siteContext.ExecuteQuery();
      //  log.Info(pubPage.ToString());


  

     //   return req.CreateResponse(web.Description);
       //  var properties = tenant.GetSitePropertiesByUrl(completeUrl, false);
       // tenant.Context.Load(properties);
      //  tenant.Context.ExecuteQueryRetry();
      //  log.Info(properties.StorageMaximumLevel.ToString());
      //  return req.CreateResponse(HttpStatusCode.OK, new { quota = properties.StorageMaximumLevel.ToString(), warning = properties.StorageWarningLevel.ToString() });

    }
    catch (Exception ex)
    {
       return req.CreateResponse(HttpStatusCode.OK, new { Exception = "Server Exception. Unable to get site details " + ex.ToString() });
    }

    log.Info("C# HTTP trigger function processed a request.");

    // parse query parameter
    string name = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "name", true) == 0)
        .Value;

    if (name == null)
    {
        // Get request body
        dynamic data = await req.Content.ReadAsAsync<object>();
        name = data?.name;
    }
    
    return name == null
        ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
        : req.CreateResponse(HttpStatusCode.OK, "Hello " + name);
        
        
}


using System;
using System.Net;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Utilities;
using PnPAuthenticationManager = OfficeDevPnP.Core.AuthenticationManager;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{

    dynamic data = await req.Content.ReadAsAsync<object>();

    string sharePointSiteUrl = data["sharePointSiteUrl"];
    string baseFolderServerRelativeUrl = data["baseFolderServerRelativeUrl"];
    string newFolderName = data["newFolderName"];

    log.Info($"sharePointSiteUrl = '{sharePointSiteUrl}'");
    log.Info($"baseFolderServerRelativeUrl = '{baseFolderServerRelativeUrl}'");
    log.Info($"newFolderName = '{newFolderName}'");

    string userName = System.Environment.GetEnvironmentVariable("SharePointUser", EnvironmentVariableTarget.Process);
    string password = System.Environment.GetEnvironmentVariable("SharePointPassword", EnvironmentVariableTarget.Process);

    var authenticationManager = new PnPAuthenticationManager();
    var clientContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(sharePointSiteUrl, userName, password);
    var pnpClientContext = PnPClientContext.ConvertFrom(clientContext);

    string newFolderUrl = UrlUtility.Combine(baseFolderServerRelativeUrl, newFolderName);

    string resultMessage = "";

    if(!doesFolderExist(pnpClientContext, newFolderUrl)){
        var folder = pnpClientContext.Web.GetFolderByServerRelativeUrl(baseFolderServerRelativeUrl);
        folder.AddSubFolder(newFolderName);

        pnpClientContext.ExecuteQuery();

        resultMessage = "Folder has been created";
    } else {
        resultMessage = "Folder already exists";
    }
    
    return req.CreateResponse(HttpStatusCode.OK, resultMessage);
}

public static bool doesFolderExist(PnPClientContext clientContext, string folderUrl)
{
    try
    {
        var folder = clientContext.Web.GetFolderByServerRelativeUrl(folderUrl);
        clientContext.ExecuteQuery();

        return true;
    }
    catch
    {
        return false;
    }
}
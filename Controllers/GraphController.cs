using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Net;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace TestMvc1.Controllers
{
    public class GraphController : Controller
    {
        public static string clientId = "";
        public static string clientSecret = "";
        public static string tenantId = "";
        public static string uriString = "http://localhost:{port number}/Graph/GetToken";
        //public static string uriString = "https://{youwebapp}.azurewebsites.net/Graph/GetToken";
        public static string authorityURL = @"https://login.microsoftonline.com/" + tenantId;
        public static string resource = @"https://graph.microsoft.com";
        

        public ActionResult Index(string authenticationCode)
        {
            string code = (string)Session["code"];
            GraphServiceClient graphClient = GetGraphClient(code);

            //Get User information [me]
            Task<User> meRequest = graphClient.Me.Request().GetAsync();
            meRequest.Wait();
            User resultMeRequest = meRequest.Result;

            Response.Write(resultMeRequest.AboutMe);

            return View();
        }

        public ActionResult GetToken()
        {
            //Necessary when behind a proxy
            //System.Net.WebRequest.DefaultWebProxy = GetDefaultProxy();

            AuthenticationContext authContext = new AuthenticationContext(authorityURL, true);

            if (Request.Params["code"] != null)
            {
                //Step 2: Using the Authorization Code from Step 1 - Request for Access Token
                string code = Request.Params["code"];

                ClientCredential clientCredentials = new ClientCredential(clientId, clientSecret);

                Task<AuthenticationResult> request = authContext.AcquireTokenByAuthorizationCodeAsync(code, new Uri(uriString), clientCredentials);
                request.Wait();

                Session["code"] = request.Result.AccessToken;
                return RedirectToAction("Index");
            }
            else
            {
                //Step 1: Get Authorization Code
                Task<Uri> redirectUri = authContext.GetAuthorizationRequestUrlAsync(resource, clientId, new Uri(uriString), UserIdentifier.AnyUser, string.Empty);
                redirectUri.Wait();
                return Redirect(redirectUri.Result.AbsoluteUri);
            }

        }

        public GraphServiceClient GetGraphClient(string graphToken)
        {
            DelegateAuthenticationProvider authenticationProvider = new DelegateAuthenticationProvider(
                (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", graphToken);
                    return Task.FromResult(0);
                });
            return new GraphServiceClient(authenticationProvider);
        }

        public IWebProxy GetDefaultProxy()
        {
            try
            {
                string DestinationUrl = "http://www.google.com";
                string PacUri = string.Empty;
                PacUri = "Your Pac Url Adress";

                // Create test request 
                WebRequest TestRequest = WebRequest.Create(DestinationUrl);

                // Optain Proxy address for the URL 
                string ProxyAddresForUrl = Proxy.GetProxyForUrlUsingPac(DestinationUrl, PacUri);
                if (ProxyAddresForUrl != null)
                {
                    if (ProxyAddresForUrl.Contains(";"))
                    {
                        string[] proxies = ProxyAddresForUrl.Split(';');
                        foreach (string proxy in proxies)
                        {
                            Console.WriteLine("Found Proxy: {0}", proxy);
                            return new WebProxy(proxy);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Found Proxy: {0}", ProxyAddresForUrl);
                        return new WebProxy(ProxyAddresForUrl);
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception("Exception occured while getting default proxy " + ex.ToString());
            }
            return null;
        }

    }
}
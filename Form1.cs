using Azure.Storage.Blobs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.ServiceModel.Description;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Manual_Notes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _toClose = true;
        }




        private IOrganizationService _serviceProxy;

        public static IOrganizationService ConnectToMSCRM()
        {
            string clientId = Convert.ToString(ConfigurationManager.AppSettings["clid"]);
            string clientSecret = Convert.ToString(ConfigurationManager.AppSettings["secret"]);
            string organizationUri = Convert.ToString(ConfigurationManager.AppSettings["orgurl"]);
            try
            {
                var conn = new CrmServiceClient($@"AuthType=ClientSecret;url={organizationUri};ClientId={clientId};ClientSecret={clientSecret}");

                return conn.OrganizationWebProxyClient != null ? conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while connecting to CRM " + ex.Message);
                return null;
            }
        }

        //public OrganizationServiceProxy ConnectToMSCRM()
        //{
        //    try
        //    {
        //        string SvrUrl = Convert.ToString(ConfigurationManager.AppSettings["SvrUrl"]);
        //        string Username = Convert.ToString(ConfigurationManager.AppSettings["Username"]);
        //        string Password = Convert.ToString(ConfigurationManager.AppSettings["Password"]);

        //        ClientCredentials credentials = new ClientCredentials();
        //        credentials.UserName.UserName = Username;
        //        credentials.UserName.Password = Password;
        //        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        //        Uri serviceUri = new Uri(SvrUrl);
        //        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        //        OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
        //        proxy.EnableProxyTypes();
        //        return proxy;
        //    }
        //    catch (Exception ex)
        //    {
        //        return null;
        //    }
        //}


        public int MaxProgress { get; set; }



        //  private void bgwAssetUpdate_DoWork(object sender, DoWorkEventArgs e)
        //  {

        //      var max = int.Parse(txtSheet.Text);
        //      string _lofFile = @"LogFile" + DateTime.Now.ToString("ddMMyyyyHHmmsstt") + ".txt";
        //      string fetchStr = @"<fetch distinct='false' mapping='logical' count='10'>
        //                              <entity name='annotation' >
        //                              <attribute name='annotationid' />
        //                              <attribute name='mimetype' />
        //                              <attribute name='filename' />
        //                              <attribute name='documentbody' />
        //                              <attribute name='isdocument' />
        //                              <attribute name='objecttypecode' />
        //                              <attribute name='objectid' />
        //                              <filter type='and'>
        //                                <condition attribute='isdocument' operator='eq' value='1' />
        //                                <condition attribute='documentbody' operator='not-null' />
        //<condition attribute='objecttypecode' operator='in'>" + getFilter() + @"</condition>
        //                              </filter>
        //                             </entity>
        //                         </fetch>";

        //      for (int i = 1; i <= max; i++)
        //      {
        //          try
        //          {
        //              _serviceProxy = ConnectToMSCRM();

        //              if (_serviceProxy != null)
        //              {
        //                  var collection = _serviceProxy.RetrieveMultiple(new FetchExpression(fetchStr));

        //                  MaxProgress = max;// collection.Entities.Count;

        //                  //DataTable dtLoop = (DataTable)dgvImport.DataSource;
        //                  if (collection.Entities.Count <= 0)
        //                  {
        //                      max = i;
        //                      bgwAssetUpdate.ReportProgress(100);
        //                      bgwAssetUpdate.CancelAsync();
        //                  }

        //                  foreach (var _ent in collection.Entities)
        //                  {
        //                      if (_serviceProxy == null) _serviceProxy = ConnectToMSCRM();
        //                      UpdateAssetsTemp(_ent, _serviceProxy, _lofFile);
        //                      if (progressBar1.Value < progressBar1.Maximum)
        //                      {
        //                          // progressBar1.Value = progressBar1.Value + 1;
        //                          bgwAssetUpdate.ReportProgress(progressBar1.Value + (1 / 10));
        //                      }


        //                  }

        //              }
        //          }
        //          catch (Exception ex)
        //          {
        //              File.AppendAllText(_lofFile, "Error:" + ex.Message + System.Environment.NewLine);
        //          }
        //      }

        //  }


        private void bgwAssetUpdate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Maximum = MaxProgress;
            progressBar1.Value = e.ProgressPercentage;
            txtProgressVal.Text = e.ProgressPercentage.ToString();
            txtProgMax.Text = MaxProgress.ToString();
        }
        public bool _toClose = false;

        private void bgwAssetUpdate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            _toClose = true;
            this.Close();
        }



        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_toClose == false)
            {

                switch (e.CloseReason)
                {
                    case CloseReason.UserClosing:
                        e.Cancel = true;
                        break;
                }
            }

            base.OnFormClosing(e);
        }
        private void btnView_Click(object sender, EventArgs e)
        {



            //if (MessageBox.Show("Do you want to Continue Move Notes ?", "Move Notes", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
            //{
            //    return;
            //}
            //_serviceProxy = ConnectToMSCRM();

            //progressBar1.Maximum = 100;
            //txtProgMax.Text = "100";
            //button1.Enabled = false;

            //bgwAssetUpdate.RunWorkerAsync();

        }

        private void UpdateAssetsTemp(Entity entity, IOrganizationService service, string _logFile)
        {

            string strRemarks = "";

            try
            {


                //MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(entity.Attributes["documentbody"].ToString()));

                var contentBytes = Convert.FromBase64String(entity.Attributes["documentbody"].ToString());
                string downloadURL = "";
                var azureUploadStat = false;
                // var azureUploadStat = UploadStream(memoryStream, entity.Attributes["filename"].ToString(), entity.Attributes["mimetype"].ToString(),out downloadURL); // UploadStreamToAzureContainer(memoryStream, entity.Attributes["filename"].ToString(), entity.Attributes["mimetype"].ToString());//, out downloadURL);
                string storageKey = "2sjBJpAAfXw6pQNYd/QUSafR/OueteMusk/gvR/+jQlqvs8Cte0fXpMAnaVVPLP7iubRJqmwTFf48XEugfxwsw==";// "vtGTF2lbLESHGtQxAK06mM4AM7SrQmCcRTVsz/5kUj8zze60obs+y6lTMvMpXix2BgTjuTdfEoSP6ilZsurosA==";// "2sjBJpAAfXw6pQNYd/QUSafR/OueteMusk/gvR/+jQlqvs8Cte0fXpMAnaVVPLP7iubRJqmwTFf48XEugfxwsw==";
                string storageAccount = "blobstoragecrm";// "dynamiccrmelgi";

                // var orgName = service.ConnectedOrgPublishedEndpoints[Microsoft.Xrm.Sdk.Discovery.EndpointType.WebApplication];
                //  service
                // context.OrganizationId.
                //orgdc21d191

                string containerName = "elgicrmfilesprod";// "elgicrmportaldev";
                //if (context != null && context.OrganizationName != null && context.OrganizationName != "" && context.OrganizationName == "orgdc21d191")
                //{
                //    containerName = "elgicrmportalprod";
                //}


                //  string blobName = entity.Attributes["filename"].ToString();
                string filename = entity.Attributes["filename"].ToString();
                var _extendedFilename = entity.Id.ToString();
                // int contentLength = Convert.FromBase64String(entity.Attributes["documentbody"].ToString()).Length;

                if (entity.Attributes.Contains("objecttypecode"))
                    _extendedFilename = _extendedFilename + "-" + entity.Attributes["objecttypecode"].ToString();

                if (entity.Attributes.Contains("objectid"))
                {
                    _extendedFilename = _extendedFilename + "-" + ((EntityReference)entity.Attributes["objectid"]).Id.ToString();

                }

                if (entity.Attributes.Contains("filename"))
                    _extendedFilename = _extendedFilename + "-SPTFM" + entity.Attributes["filename"].ToString();

                _extendedFilename = _extendedFilename.Replace(" ", "").ToLower().Replace("-", "");

                filename = _extendedFilename;


                filename = Regex.Replace(filename, @"[^0-9a-zA-Z]+", ".");

                string requestUri = @"https://" + storageAccount + ".blob.core.windows.net/" + containerName + "/" + filename;

                // string content = Encoding.UTF8.GetString(Convert.FromBase64String(entity.Attributes["documentbody"].ToString()));
                //int contentLength = Encoding.UTF8.GetByteCount(content);

                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Put, requestUri);
                httpRequestMessage.Content = new ByteArrayContent(contentBytes);  //(HttpContent)Encoding.UTF8.GetBytes(content);// new ByteArrayContent(Convert.FromBase64String(entity.Attributes["documentbody"].ToString()));

                DateTime now = DateTime.UtcNow;
                httpRequestMessage.Headers.Add("x-ms-date", now.ToString("R", CultureInfo.InvariantCulture));
                httpRequestMessage.Headers.Add("x-ms-version", "2018-11-09");
                httpRequestMessage.Headers.Add("x-ms-blob-type", "BlockBlob");

                httpRequestMessage.Headers.Authorization = GetAuthorizationHeader(
                storageAccount, storageKey, now, httpRequestMessage);


                using (HttpResponseMessage httpResponseMessage = new HttpClient().SendAsync(httpRequestMessage).Result)
                {
                    //   tracingService.Trace("Send httpResponseMessage : Status Code " + httpResponseMessage.StatusCode.ToString());
                    // If successful (status code = 200), 
                    //   parse the XML response for the container names.
                    if (httpResponseMessage.StatusCode == HttpStatusCode.Created)
                    {
                        //String xmlString =  httpResponseMessage.Content.ReadAsStringAsync().Result;
                        azureUploadStat = true;
                        downloadURL = requestUri;
                    }
                    else
                    {
                        //tracingService.Trace("Excep tion 1 ");
                        throw new InvalidPluginExecutionException(OperationStatus.Failed, "Azure http Response - " + httpResponseMessage.StatusCode.ToString());
                    }
                }


                if (azureUploadStat == true && downloadURL != "")
                {
                    Entity entityUpd = new Entity(entity.LogicalName, entity.Id);

                    // tracingService.Trace("Url " + downloadURL);

                    string _subj = "";
                    if (entity.Attributes.Contains("subject"))
                    {
                        _subj = entity.Attributes["subject"].ToString();
                    }

                    entityUpd.Attributes["subject"] = "[" + _subj + "] URL[" + downloadURL + "]";
                    entityUpd.Attributes["isdocument"] = false;
                    // entity.Attributes.Remove("documentbody");
                    // entity.Attributes.Remove("filename");
                    // entity.Attributes.Remove("mimetype");
                    entityUpd.Attributes["documentbody"] = null;
                    entityUpd.Attributes["filename"] = null;
                    entityUpd.Attributes["mimetype"] = null;

                    service.Update(entityUpd);

                    File.AppendAllText(_logFile, downloadURL + System.Environment.NewLine);

                    //Entity e = new Entity("annotation");
                    //e.Id = entity.Id;
                    //e.Attributes["subject"] = downloadURL;
                    //e.Attributes["isdocument"] = false;
                    //e.Attributes["documentbody"] = null;
                    //e.Attributes["filename"] = null;
                    //e.Attributes["mimetype"] = null;
                    //service.Update(e);

                }
                else
                {

                    throw new InvalidPluginExecutionException(OperationStatus.Failed, "Uploading to Azure Blob Storage Failed");
                }
            }
            catch (Exception e)
            {

                //throw new InvalidPluginExecutionException(OperationStatus.Failed, e.Message);
            }



        }

        public AuthenticationHeaderValue GetAuthorizationHeader(
                  string storageAccountName, string storageAccountKey, DateTime now,
                  HttpRequestMessage httpRequestMessage, string ifMatch = "", string md5 = "")
        {
            var contentLength = httpRequestMessage.Content.Headers.ContentLength.ToString();
            //contentLength = "4";

            HttpMethod method = httpRequestMessage.Method;
            String MessageSignature = String.Format("{0}\n\n\n{1}\n{5}\n\n\n\n{2}\n\n\n\n{3}{4}",
                      method.ToString(),
                      (method == HttpMethod.Get || method == HttpMethod.Head) ? String.Empty
                        : contentLength,
                      ifMatch,
                      GetCanonicalizedHeaders(httpRequestMessage),
                      GetCanonicalizedResource(httpRequestMessage.RequestUri, storageAccountName),
                      md5);

            // Now turn it into a byte array.
            byte[] SignatureBytes = Encoding.UTF8.GetBytes(MessageSignature);

            // Create the HMACSHA256 version of the storage key.
            HMACSHA256 SHA256 = new HMACSHA256(Convert.FromBase64String(storageAccountKey));

            // Compute the hash of the SignatureBytes and convert it to a base64 string.
            string signature = Convert.ToBase64String(SHA256.ComputeHash(SignatureBytes));

            // This is the actual header that will be added to the list of request headers.
            // You can stop the code here and look at the value of 'authHV' before it is returned.
            AuthenticationHeaderValue authHV = new AuthenticationHeaderValue("SharedKey",
                storageAccountName + ":" + Convert.ToBase64String(SHA256.ComputeHash(SignatureBytes)));
            return authHV;
        }
        public string GetCanonicalizedHeaders(HttpRequestMessage httpRequestMessage)
        {
            var headers = from kvp in httpRequestMessage.Headers
                          where kvp.Key.StartsWith("x-ms-", StringComparison.OrdinalIgnoreCase)
                          orderby kvp.Key
                          select new { Key = kvp.Key.ToLowerInvariant(), kvp.Value };

            StringBuilder sb = new StringBuilder();

            // Create the string in the right format; this is what makes the headers "canonicalized" --

            foreach (var kvp in headers)
            {
                StringBuilder headerBuilder = new StringBuilder(kvp.Key);
                char separator = ':';

                // Get the value for each header, strip out \r\n if found, then append it with the key.
                foreach (string headerValues in kvp.Value)
                {
                    string trimmedValue = headerValues.TrimStart().Replace("\r\n", String.Empty);
                    headerBuilder.Append(separator).Append(trimmedValue);

                    // Set this to a comma; this will only be used 
                    //   if there are multiple values for one of the headers.
                    separator = ',';
                }
                sb.Append(headerBuilder.ToString()).Append("\n");
            }
            return sb.ToString();
        }
        public string GetCanonicalizedResource(Uri address, string storageAccountName)
        {
            // The absolute path is "/" because for we're getting a list of containers.
            StringBuilder sb = new StringBuilder("/").Append(storageAccountName).Append(address.AbsolutePath);

            // Address.Query is the resource, such as "?comp=list".
            // This ends up with a NameValueCollection with 1 entry having key=comp, value=list.
            // It will have more entries if you have more query parameters.

            //NameValueCollection values = ParseQueryString(address.Query);

            //foreach (var item in values.AllKeys.OrderBy(k => k))
            //{
            //    sb.Append('\n').Append(item).Append(':').Append(values[item]);
            //}

            return sb.ToString();

        }

        public static NameValueCollection ParseQueryString(string s)
        {
            NameValueCollection nvc = new NameValueCollection();
            // remove anything other than query string from url
            if (s.Contains("?"))
            {
                s = s.Substring(s.IndexOf('?') + 1);
            }
            foreach (string vp in Regex.Split(s, "&"))
            {
                string[] singlePair = Regex.Split(vp, "=");
                if (singlePair.Length == 2)
                {
                    nvc.Add(singlePair[0], singlePair[1]);
                }
                else
                {
                    // only one key with no value specified in query string
                    nvc.Add(singlePair[0], string.Empty);
                }
            }
            return nvc;
        }

        private void button1_Click(object sender, EventArgs e)
        {


        }

        public string getFilter()
        {
            string[] temp = { "account", "contact", "appointment", "email", "quote", "opportunity", "salesorder", "msdyn_customerasset", "elgi_partmaster", "elgi_monthlyplan", "elgi_focapprovalmatrix", "elgi_foc", "elgi_feedback", "elgi_dealerworkorder", "elgi_dealerservicereimbursement", "elgi_customervisitfab", "elgi_dealerofferproduct", "elgi_dealeroffer", "elgi_dealerindentline", "elgi_dealerindent", "elgi_dealerbulkcv", "elgi_ddc", "msdyn_workorder", "msdyn_workorderproduct", "msdyn_workorderservice", "msdyn_workorderservicetask", "msdyn_workorderincident", "msdyn_incidenttype" };
            string _out = "";
            var multipleRequest = new ExecuteMultipleRequest()
            {

                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };

            foreach (string s in temp)
            {
                RetrieveEntityRequest _eRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = s
                };
                // RetrieveEntityResponse _eResponse = (RetrieveEntityResponse)_serviceProxy.Execute(_eRequest);
                multipleRequest.Requests.Add(_eRequest);

            }





            //foreach (var entity in entities)
            //{
            //    UpdateRequest updateRequest = new UpdateRequest { Target = entity };
            //    multipleRequest.Requests.Add(updateRequest);
            //}

            // Execute all the requests in the request collection using a single web method call.
            if (_serviceProxy == null) _serviceProxy = ConnectToMSCRM();
            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)_serviceProxy.Execute(multipleRequest);

            foreach (ExecuteMultipleResponseItem _itemObj in multipleResponse.Responses)
            {
                var _eResponse = (RetrieveEntityResponse)_itemObj.Response;
                _out = _out + "<value>" + _eResponse.EntityMetadata.ObjectTypeCode.Value + "</value>";
            }


            return _out;
        }
        private void UpdateAttachment(byte[] contentBytes, string filename, string _logFile, string _erFile)
        {



            try
            {




                //var contentBytes = Convert.FromBase64String(entity.Attributes["documentbody"].ToString());
                string downloadURL = "";
                var azureUploadStat = false;

                string storageKey = "2sjBJpAAfXw6pQNYd/QUSafR/OueteMusk/gvR/+jQlqvs8Cte0fXpMAnaVVPLP7iubRJqmwTFf48XEugfxwsw==";// 
                string storageAccount = "blobstoragecrm";// "dynamiccrmelgi";


                string containerName = "elgicrmfilesprod";// "elgicrmportaldev";


                string requestUri = @"https://" + storageAccount + ".blob.core.windows.net/" + containerName + "/" + filename;

                // string content = Encoding.UTF8.GetString(Convert.FromBase64String(entity.Attributes["documentbody"].ToString()));
                //int contentLength = Encoding.UTF8.GetByteCount(content);

                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Put, requestUri);
                httpRequestMessage.Content = new ByteArrayContent(contentBytes);  //(HttpContent)Encoding.UTF8.GetBytes(content);// new ByteArrayContent(Convert.FromBase64String(entity.Attributes["documentbody"].ToString()));

                DateTime now = DateTime.UtcNow;
                httpRequestMessage.Headers.Add("x-ms-date", now.ToString("R", CultureInfo.InvariantCulture));
                httpRequestMessage.Headers.Add("x-ms-version", "2018-11-09");
                httpRequestMessage.Headers.Add("x-ms-blob-type", "BlockBlob");

                httpRequestMessage.Headers.Authorization = GetAuthorizationHeader(
                storageAccount, storageKey, now, httpRequestMessage);


                using (HttpResponseMessage httpResponseMessage = new HttpClient().SendAsync(httpRequestMessage).Result)
                {
                    //   tracingService.Trace("Send httpResponseMessage : Status Code " + httpResponseMessage.StatusCode.ToString());
                    // If successful (status code = 200), 
                    //   parse the XML response for the container names.
                    if (httpResponseMessage.StatusCode == HttpStatusCode.Created)
                    {
                        //String xmlString =  httpResponseMessage.Content.ReadAsStringAsync().Result;
                        azureUploadStat = true;
                        downloadURL = requestUri;
                    }
                    else
                    {
                        //tracingService.Trace("Excep tion 1 ");
                        throw new InvalidPluginExecutionException(OperationStatus.Failed, "Azure http Response - " + httpResponseMessage.StatusCode.ToString());
                    }
                }


                if (azureUploadStat == true && downloadURL != "")
                {


                    File.AppendAllText(_logFile, downloadURL + System.Environment.NewLine);


                }
                else
                {
                    File.AppendAllText(_erFile, filename + System.Environment.NewLine);
                    //throw new InvalidPluginExecutionException(OperationStatus.Failed, "Uploading to Azure Blob Storage Failed");
                }
            }
            catch (Exception e)
            {
                File.AppendAllText(_erFile, filename + System.Environment.NewLine);
                //throw new InvalidPluginExecutionException(OperationStatus.Failed, e.Message);
            }



        }

        public static void BulkDelete(IOrganizationService service, DataCollection<Entity> entityReferences, string logFile)
        {

            // Create an ExecuteMultipleRequest object.
            var multipleRequest = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses. 
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };
            //string _log = "";
            // Add a DeleteRequest for each entity to the request collection.
            foreach (var entityRef in entityReferences)
            {
                //log = entityRef )
                // _log = JsonConvert.SerializeObject(entityRef);
                //File.AppendAllText(logFile, _log + System.Environment.NewLine);
                DeleteRequest deleteRequest = new DeleteRequest { Target = entityRef.ToEntityReference() };
                multipleRequest.Requests.Add(deleteRequest);
            }

            // Execute all the requests in the request collection using a single web method call.
            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);

        }
        public static void MoveToAzure(DataCollection<EntityReference> entityReferences)
        {

        }

        private static ManualResetEvent mre = new ManualResetEvent(false);

        private void ProcessFiles()
        {

            string _onorbefore = "2023-08-31";
            int _bulkdelCount = 50;

            MethodInvoker action1122 = delegate
                {
                //textBox1.Text = "Connected to server... \n";
                txtOnOfBefore.Text = _onorbefore;
                };
            txtOnOfBefore.BeginInvoke(action1122);


            // string logFile = @"CheckListItemsLogFile" + DateTime.Now.ToString("ddMMyyyyHHmmsstt") + ".txt";
            string logFile = @"CheckListItemsCountLogFile" + ".txt";


            var fetch = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' count='" + _bulkdelCount + @"' ##PAGE##>
              <entity name='msdyn_workorderservicetask'>
                <attribute name='msdyn_workorderservicetaskid' />
                <filter type='and'>
	              <condition attribute='createdon' operator='on-or-before' value='" + _onorbefore + @"' />
                  <condition attribute='msdyn_workorder' operator='not-null' />
                  <condition attribute='msdyn_workorderincident' operator='not-null' />
                </filter>
              </entity>
            </fetch>
".Trim(); // PAGE SIZE = 100



            // ADD FETCHXML COUNT 200
            var i = 1;


            var moreRecords = false;

            int page = 1;

            var cookie = string.Empty;

            // List<Entity> Entities = new List<Entity>();

            do

            {
                if (_serviceProxy == null)
                {
                    _serviceProxy = ConnectToMSCRM();
                }

                var xml = fetch.Replace("##PAGE##", cookie);

                var collection = _serviceProxy.RetrieveMultiple(new FetchExpression(xml));

                if (collection.Entities.Count > 0)
                {
                    //MoveToAzure();
                    mre.WaitOne(4000);
                    try
                    {
                        BulkDelete(_serviceProxy, collection.Entities, logFile);

                        //MethodInvoker _writeLog = delegate
                        //{
                        //    //textBox1.Text = "Connected to server... \n";
                        //    //txtMax.Enabled = true;
                        File.AppendAllText(logFile, DateTime.Now.ToString("dd-MM-yyyy HH:mm") + "    " + _bulkdelCount.ToString() + System.Environment.NewLine);
                        //};
                        //txtMax.BeginInvoke(_writeLog);
                    }
                    catch (Exception ex)
                    {

                    }
                    MethodInvoker action2 = delegate
                    {
                        textBox1.Text = i.ToString();
                    };
                    textBox1.BeginInvoke(action2);
                }



                moreRecords = collection.MoreRecords && i < MaxProgress;

                if (moreRecords)

                {
                    mre.WaitOne(2000);
                    page++;

                    cookie = string.Format("paging-cookie='{0}' page='{1}'", System.Security.SecurityElement.Escape(collection.PagingCookie), page);

                }
                i++;
            } while (moreRecords);







            MethodInvoker action = delegate
            {
                //textBox1.Text = "Connected to server... \n";
                textBox1.Text = (i - 1).ToString() + " Completed";
                _toClose = true;
            };
            textBox1.BeginInvoke(action);



            MethodInvoker action1 = delegate
            {
                //textBox1.Text = "Connected to server... \n";
                button1.Enabled = true;
            };
            button1.BeginInvoke(action1);

            //MethodInvoker action15 = delegate
            //{
            //    //textBox1.Text = "Connected to server... \n";
            //    txtMax.Enabled = true;
            //};
            //txtMax.BeginInvoke(action15);






            MessageBox.Show("Completed");
        }



        private void button1_Click_1(object sender, EventArgs e)
        {


            if (_serviceProxy == null)
            {
                _serviceProxy = ConnectToMSCRM();
            }
            MaxProgress = int.Parse(txtMax.Text);
            txtMax.Enabled = false;
            button1.Enabled = false;
            _toClose = false;
            ThreadStart childref = new ThreadStart(ProcessFiles);


            //Console.WriteLine("In Main: Creating the Child thread");
            Thread childThread = new Thread(childref);
            childThread.Start();
            // Console.ReadKey();



        }
        public static List<Entity> RetrieveAllRecords(IOrganizationService service, string fetch)

        {

            var moreRecords = false;

            int page = 1;

            var cookie = string.Empty;

            List<Entity> Entities = new List<Entity>();

            do

            {

                var xml = fetch.Replace("##PAGE##", cookie);

                var collection = service.RetrieveMultiple(new FetchExpression(xml));

                if (collection.Entities.Count >= 0)

                    Entities.AddRange(collection.Entities);

                moreRecords = collection.MoreRecords;

                if (moreRecords)

                {

                    page++;

                    cookie = string.Format("paging-cookie='{0}' page='{1}'", System.Security.SecurityElement.Escape(collection.PagingCookie), page);

                }

            } while (moreRecords);

            return Entities;

        }
        private void button2_Click(object sender, EventArgs e)
        {

            return;
            _toClose = false;
            string fetchStr = @"<fetch distinct='false' mapping='logical' ##PAGE## >
                                    <entity name='annotation' >
                                    <attribute name='annotationid' />
                                    <attribute name='subject' />
                                    <filter type='and'>
                                       <condition attribute='subject' operator='like' value='%dynamiccrmelgi%' />
                                       <condition attribute='modifiedon' operator='on-or-after' value='2022-01-01' />
 <condition attribute='modifiedon' operator='on-or-before' value='2022-05-01' />
                                    </filter>
                                   </entity>
                               </fetch>";


            if (_serviceProxy == null) _serviceProxy = ConnectToMSCRM();
            //var collection = _serviceProxy.RetrieveMultiple(new FetchExpression(fetchStr));

            var collection = RetrieveAllRecords(_serviceProxy, fetchStr);

            foreach (var _ent in collection)
            {





                // tracingService.Trace("Url " + downloadURL);
                string _olddata = @"https://dynamiccrmelgi.blob.core.windows.net/elgicrmportalprod";
                string _newdata = @"https://blobstoragecrm.blob.core.windows.net/elgicrmfilesprod";

                string _subj = "";
                if (_ent.Attributes.Contains("subject"))
                {
                    _subj = _ent.Attributes["subject"].ToString().Replace(_olddata, _newdata);

                    Entity entityUpd = new Entity(_ent.LogicalName, _ent.Id);
                    entityUpd.Attributes["subject"] = _subj;
                    _serviceProxy.Update(entityUpd);
                }






            }
            _toClose = true;
            MessageBox.Show("Completed");
        }
    }
}

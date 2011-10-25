using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using System.IO;

using IntranetDASHBOARD.API;

namespace SharePointConnector
{
    public partial class GetImage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            DownloadImageFile();
        }


        public void DownloadImageFile()
        {
            //decode query string values
            Guid key = new Guid(HttpUtility.UrlDecode(Request.QueryString["key"]));
            string hostName = HttpUtility.UrlDecode(Request.QueryString["host"]);
            string fileName = HttpUtility.UrlDecode(Request.QueryString["fileType"]);

            //retrieve data for this component
            iDCMSData componentData = Intranet.Get.GetiDCMSData(key);
            //create a network credential object to use when downloading the file
            string username = componentData.ContainsKey("Username") ? componentData["Username"] : string.Empty;
            string password = componentData.ContainsKey("Password") ? componentData["Password"] : string.Empty;
            string domain = componentData.ContainsKey("Domain") ? componentData["Domain"] : string.Empty;
            NetworkCredential sharepointCredentitals = new NetworkCredential(username, password, domain);

            Response.Clear();
            Response.AddHeader("Content-Type", "image/jpeg;");
            

            string url = hostName + "/_layouts/images/";
            try
            {
                WebClient sharepointFileClient = new WebClient();
                sharepointFileClient.Credentials = sharepointCredentitals;
                byte[] content = sharepointFileClient.DownloadData(url + "/" + fileName);

                if (content != null)
                {
                    Response.Cache.SetMaxAge(new TimeSpan(2, 0, 0, 0, 0));
                    Stream dataStream = Response.OutputStream;
                    dataStream.Write(content, 0, content.Length);
                }
                else
                {
                    DownloadImageFromiD();
                }

            }
            catch (Exception ex)
            {
                DownloadImageFromiD();
            }
            finally
            {
                Response.End();
            }

            
        }

        private void DownloadImageFromiD()
        {
            FileStream stream = new FileStream(Server.MapPath("/") + "/images/icons/filetypes_16/icon_typ16_various.gif", FileMode.Open, FileAccess.Read);
            BinaryReader reader = new BinaryReader(stream);
            Byte[] bytes = reader.ReadBytes((int)stream.Length);
            reader.Close();
            stream.Close();
            Response.BinaryWrite(bytes);
        }
    }
}
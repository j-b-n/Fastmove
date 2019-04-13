using Polenter.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace FastMove
{
    public class FastMoveUpdateInfoVariables
    {
        string _version = "";
        public string Version
        {
            get { return _version; }
            set { _version = value; }
        }
    }

    /// <summary>
    /// Check if there is an update available to the Add-In
    /// </summary>
    /// 

    class UpdateInfo
    {
        private string publishedVersion = "0.0.0.0";

        private string GetRunningVersion()
        {
            publishedVersion = "0.0.0.0";
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                System.Deployment.Application.ApplicationDeployment currDeploy = System.Deployment.Application.ApplicationDeployment.CurrentDeployment;
                Version pubVer = currDeploy.CurrentVersion;
                publishedVersion = pubVer.Major.ToString() + "." + pubVer.Minor.ToString() + "." +
                    pubVer.Build.ToString() + "." + pubVer.Revision.ToString();
                return publishedVersion;
            }            
            return publishedVersion;
        }


        public void WriteOnlineUpdateInfo()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FastMove";

            try
            {
                // If the directory doesn't exist, create it.
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }
            catch (Exception)
            {
                // Fail silently
            }

            path += "\\FastMoveOnlineVersion.xml";

            try
            {
                FastMoveUpdateInfoVariables UpdateVariables = new FastMoveUpdateInfoVariables
                {
                    Version = "0.0.0.0"
                };
                var serializer = new SharpSerializer();
                serializer.Serialize(UpdateVariables, path);
            }
            catch (Exception)
            {                
            }
        }

        public Stream GenerateStreamFromString(string s)
        {
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        /// <summary>
        /// Check GitHub for the current published version
        /// </summary>
        /// <returns>
        /// 0 - No update available
        /// 1 - Update available</returns>
        public int CheckForUpdate()
        {
            FastMoveUpdateInfoVariables UpdateVariables;
            string downloadString;            
            try
            {
                WebClient client = new WebClient();
                downloadString = client.DownloadString("https://raw.githubusercontent.com/j-b-n/Fastmove/master/update.xml");                
            }
            catch
            {
                return 0;
            }

            //writeOnlineUpdateInfo();

            var serializer = new SharpSerializer(false);
            
            using (Stream s = GenerateStreamFromString(downloadString))
            {
                if (s.Length > 0)
                {
                    try
                    {
                        UpdateVariables = (FastMoveUpdateInfoVariables)serializer.Deserialize(s);
                    }
                    catch
                    {
                        return 0;
                    }
                    string runningVersion = GetRunningVersion();

                    if (UpdateVariables.Version != runningVersion)
                    {
                        //New version available! 
                        return 1;
                    }
                }             
            }                       
            return 0;
        }
    }
}

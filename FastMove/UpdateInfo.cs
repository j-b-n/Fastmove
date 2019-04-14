using Polenter.Serialization;
using System;
using System.IO;
using System.Net;

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

        public FastMoveUpdateInfoVariables UpdateVariables = new FastMoveUpdateInfoVariables
        {
            Version = "0.0.0.0"
        };       


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

        private bool UpdateCache()
        {
            bool update = false;
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FastMove";

            try
            {
                // If the directory doesn't exist, create it.
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }
            catch
            {
                // Fail silently
            }

            path += "\\OnlineVersion.xml";

            if (File.Exists(path))
            {
                if(File.GetLastWriteTime(path).AddDays(7) > DateTime.Now)
                {
                    update = true;
                }

            }
            else
            {
                update = true;
            }

            if (update)
            {
                WebClient client = new WebClient();
                string downloadString = client.DownloadString("https://raw.githubusercontent.com/j-b-n/Fastmove/master/update.xml");

                File.WriteAllText(path, downloadString);
            }

            return update;
        }

        /// <summary>
        /// Check GitHub for the current published version
        /// </summary>
        /// <returns>
        /// 0 - No update available
        /// 1 - Update available</returns>
        public int CheckForUpdate()
        {            
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FastMove\\OnlineVersion.xml";
            
            UpdateCache();
           
            try
            {
                var serializer = new SharpSerializer(false);
                UpdateVariables = (FastMoveUpdateInfoVariables)serializer.Deserialize(path);
                
                if (UpdateVariables.Version != Globals.ThisAddIn.publishedVersion)
                {
                    //New version available! 
                    return 1;
                }
            }
            catch
            {
                return -2;
            }            
            return 0;
        }
    }
}

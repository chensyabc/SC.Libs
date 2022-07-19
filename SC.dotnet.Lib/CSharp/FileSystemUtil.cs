using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Util
{
    public class FileSystemUtil
    {
        public static void FileCopyWithFileNameAndExtension(string sourceFullPathWithExtension, string aimFileNamewithExtension)
        {
            File.Copy(sourceFullPathWithExtension, aimFileNamewithExtension, true);
        }

        public static void FileContentReplace(string fullPathWithExtension, List<KeyValuePair<string, string>> replaceDictionary)
        {
            string wholeContent = "";

            using (StreamReader reader = File.OpenText(fullPathWithExtension))
            {
                wholeContent = reader.ReadToEnd();
            }

            foreach (KeyValuePair<string, string> replaceItem in replaceDictionary)
            {
                wholeContent = wholeContent.Replace(replaceItem.Key, replaceItem.Value);
            }

            File.WriteAllText(fullPathWithExtension, wholeContent);
        }

        public static bool OpenFolder(string folderFullPath, ref string result)
        {
            if (Directory.Exists(folderFullPath))
            {
                Process.Start(folderFullPath);

                return true;
            }
            else
            {
                result = string.Format("Folder '{0}' doesn't exist.", folderFullPath);

                return false;
            }
        }

        public static bool OpenFile(string fileFullPath, ref string result)
        {
            bool isSuccessful = false;

            try
            {
                if (File.Exists(fileFullPath))
                {
                    Process.Start(fileFullPath);

                    isSuccessful = true;
                }
                else
                {
                    result = string.Format("File '{0}' doesn't exist.", fileFullPath);

                    isSuccessful = false;
                }
            }
            catch (Win32Exception w32ex)
            {
                if (w32ex.Message == "The operation was canceled by the user")
                    isSuccessful = true;
                else
                    throw;
            }
            catch (Exception)
            {
                throw;
            }

            return isSuccessful;
        }

        public static FileSystemWatcher InitialFileWatcher(string pattern, string folderFullPath, FileSystemEventHandler fileCreatedHandler)
        {
            FileSystemWatcher fileWatcher = new FileSystemWatcher()
            {
                Path = folderFullPath,
                IncludeSubdirectories = false,
                Filter = pattern,
                EnableRaisingEvents = true,
            };

            fileWatcher.Created += fileCreatedHandler;

            return fileWatcher;
        }

        public static FileSystemWatcher InitialFileWatcher(string pattern, string upLevelFolderPath, string fileParentFolder, FileSystemEventHandler fileCreatedHandler)
        {
            string folderFullPath = Path.Combine(upLevelFolderPath, fileParentFolder);

            return FileSystemUtil.InitialFileWatcher(pattern, folderFullPath, fileCreatedHandler);
        }

        public static FileInfo GetLatestFileInPathWithPattern(string serverFolderPath, string pattern)
        {
            DirectoryInfo info = new DirectoryInfo(serverFolderPath);

            var lastFileInfo = info.GetFiles(pattern, SearchOption.TopDirectoryOnly).OrderBy(f => f.Name).LastOrDefault();

            return lastFileInfo;
        }

        public static bool CreateEmptyFileWithExtension(string folderPath, string fileNameWithExtension)
        {
            FileStream newFile = File.Create(Path.Combine(folderPath, fileNameWithExtension));

            newFile.Close();

            return true;
        }
    }
}

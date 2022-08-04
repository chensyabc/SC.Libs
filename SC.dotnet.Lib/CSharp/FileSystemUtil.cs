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

        public static bool Open(string filePath, ref string message)
        {
            try
            {
                Open(filePath);
                return true;
            }
            catch (PlatformNotSupportedException)
            {
                message = "Your operating system does not support opening this file.";
            }
            catch (Exception ex)
            {
                message = "Error occurred, could not open the generated file.";
            }

            return false;
        }

        public static Process Open(string filePath)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true,
                Verb = "open",
            };

            return Process.Start(startInfo);
        }


        #region Judge File Extension

        public static bool IsPdf(string file)
        {
            bool isPdf = false;
            if (!string.IsNullOrEmpty(file) && File.Exists(file))
            {
                FileInfo fileInfo = new FileInfo(file);
                if (fileInfo.Extension.ToLower().Contains("pdf"))
                {
                    isPdf = true;
                }
            }
            return isPdf;
        }

        public static bool IsExcel(string file)
        {
            bool isExcel = false;
            if (!string.IsNullOrEmpty(file) && File.Exists(file))
            {
                FileInfo fileInfo = new FileInfo(file);
                if (fileInfo.Extension.ToLower().Contains("xls") || fileInfo.Extension.ToLower().Contains("xlsx") || fileInfo.Extension.ToLower().Contains("xlsm"))
                {
                    isExcel = true;
                }
            }
            return isExcel;
        }

        public static bool IsWord(string file)
        {
            bool isDoc = false;
            if (!string.IsNullOrEmpty(file) && File.Exists(file))
            {
                FileInfo fileInfo = new FileInfo(file);
                if (fileInfo.Extension.ToLower().Contains("doc") || fileInfo.Extension.ToLower().Contains("docx"))
                {
                    isDoc = true;
                }
            }
            return isDoc;
        }

        public static bool IsPPT(string file)
        {
            bool isPPT = false;
            if (!string.IsNullOrEmpty(file) && File.Exists(file))
            {
                FileInfo fileInfo = new FileInfo(file);
                if (fileInfo.Extension.ToLower().Contains("ppt") || fileInfo.Extension.ToLower().Contains("pptx") || fileInfo.Extension.ToLower().Contains("pptm"))
                {
                    isPPT = true;
                }
            }
            return isPPT;
        }

        public static bool IsTxt(string file)
        {
            bool isPdf = false;
            if (!string.IsNullOrEmpty(file) && File.Exists(file))
            {
                FileInfo fileInfo = new FileInfo(file);
                if (fileInfo.Extension.ToLower().Contains("txt"))
                {
                    isPdf = true;
                }
            }
            return isPdf;
        }

        public static bool IsImage(string file)
        {
            bool isImage = false;
            if (!string.IsNullOrEmpty(file) && File.Exists(file))
            {
                FileInfo fileInfo = new FileInfo(file);
                if (fileInfo.Extension.ToLower().Contains("gif") ||
                    fileInfo.Extension.ToLower().Contains("jpeg") || fileInfo.Extension.ToLower().Contains("jpg") ||
                    fileInfo.Extension.ToLower().Contains("tig") || fileInfo.Extension.ToLower().Contains("psd") ||
                    fileInfo.Extension.ToLower().Contains("raw") || fileInfo.Extension.ToLower().Contains("png"))
                {
                    isImage = true;
                }
            }
            return isImage;
        }

        #endregion


        #region Check if file in using

        public static bool IsFileInUse(FileInfo file)
        {
            FileStream stream = null;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }
            return false;
        }

        public static bool IsFileInUse(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            return IsFileInUse(file);
        }

        #endregion
    }
}

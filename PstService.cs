using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PstMerger
{
    public class PstService
    {
        public void MergeFiles(string[] sourceFiles, string destinationPst, System.Threading.CancellationToken ct, Action<int, string> onProgress)
        {
            Outlook.Application outlookApp = null;
            Outlook.NameSpace ns = null;
            Outlook.Folder destRoot = null;

            try
            {
                outlookApp = new Outlook.Application();
                ns = outlookApp.GetNamespace("MAPI");

                // 1. Ensure the destination PST exists or create it
                if (!File.Exists(destinationPst))
                {
                    onProgress(0, "Creating destination PST...");
                    ns.AddStore(destinationPst);
                }
                else
                {
                    onProgress(0, "Opening existing destination PST...");
                    ns.AddStore(destinationPst);
                }

                // Get the destination root folder
                destRoot = GetRootFolder(ns, destinationPst, onProgress);
                if (destRoot == null) throw new Exception("Could not find destination root.");

                int count = 0;
                foreach (string sourceFile in sourceFiles)
                {
                    if (ct.IsCancellationRequested) break;
                    
                    // Skip if it's the destination itself
                    if (string.Equals(Path.GetFullPath(sourceFile), Path.GetFullPath(destinationPst), StringComparison.OrdinalIgnoreCase))
                        continue;

                    count++;
                    onProgress(count, string.Format("Merging: {0}", Path.GetFileName(sourceFile)));

                    ProcessSourcePst(ns, sourceFile, destRoot, ct, onProgress);
                }

                ns.RemoveStore(destRoot);
            }
            finally
            {
                if (ns != null) Marshal.ReleaseComObject(ns);
                if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
            }
        }

        private void ProcessSourcePst(Outlook.NameSpace ns, string filePath, Outlook.Folder destRoot, System.Threading.CancellationToken ct, Action<int, string> onProgress)
        {
            Outlook.Folder sourceRoot = null;
            try
            {
                ns.AddStore(filePath);
                sourceRoot = GetRootFolder(ns, filePath, onProgress);
                if (sourceRoot == null) return;

                CopyFolders(sourceRoot, destRoot, ct, onProgress);

                ns.RemoveStore(sourceRoot);
                Marshal.ReleaseComObject(sourceRoot);
            }
            catch (Exception ex)
            {
                onProgress(-1, string.Format("Error processing {0}: {1}", Path.GetFileName(filePath), ex.Message));
            }
        }

        private void CopyFolders(Outlook.Folder sourceFolder, Outlook.Folder destFolder, System.Threading.CancellationToken ct, Action<int, string> onProgress)
        {
            if (ct.IsCancellationRequested) return;

            // 1. Copy items in the current folder
            Outlook.Items sourceItems = sourceFolder.Items;
            int itemCount = sourceItems.Count;
            
            for (int i = itemCount; i >= 1; i--)
            {
                if (ct.IsCancellationRequested) break;

                object item = null;
                dynamic copy = null;
                try
                {
                    item = sourceItems[i];
                    
                    // We copy and then move to preserve the source PST in case of failure
                    // Use dynamic to call Copy/Move on any Outlook item type
                    dynamic dynItem = item;
                    copy = dynItem.Copy();
                    copy.Move(destFolder);
                }
                catch (Exception ex)
                {
                    onProgress(-1, string.Format("Warning: Failed to copy item in {0}: {1}", sourceFolder.Name, ex.Message));
                }
                finally
                {
                    if (copy != null) Marshal.ReleaseComObject(copy);
                    if (item != null) Marshal.ReleaseComObject(item);
                }
            }
            if (sourceItems != null) Marshal.ReleaseComObject(sourceItems);

            // 2. Recursively process subfolders
            Outlook.Folders sourceSubFolders = sourceFolder.Folders;
            foreach (Outlook.Folder sourceSubFolder in sourceSubFolders)
            {
                if (ct.IsCancellationRequested) break;

                Outlook.Folder destSubFolder = null;
                Outlook.Folders destFolders = destFolder.Folders;
                
                // Try to find if subfolder exists in destination, with retry for transient COM errors
                int maxRetries = 3;
                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    try
                    {
                        destSubFolder = FindFolderByName(destFolders, sourceSubFolder.Name);
                        
                        if (destSubFolder == null)
                        {
                            try
                            {
                                destSubFolder = destFolders.Add(sourceSubFolder.Name, sourceSubFolder.DefaultItemType) as Outlook.Folder;
                            }
                            catch
                            {
                                // Fallback: Try adding without type (sometimes needed for Root folders or special stores)
                                destSubFolder = destFolders.Add(sourceSubFolder.Name) as Outlook.Folder;
                            }
                        }
                        break; // Success, exit retry loop
                    }
                    catch (Exception ex)
                    {
                        if (attempt == maxRetries)
                        {
                            onProgress(-1, string.Format("Error creating folder {0} after {1} attempts: {2}", sourceSubFolder.Name, maxRetries, ex.Message));
                        }
                        else
                        {
                            onProgress(-1, string.Format("Retry {0}/{1} for folder {2}: {3}", attempt, maxRetries, sourceSubFolder.Name, ex.Message));
                            System.Threading.Thread.Sleep(500);
                        }
                    }
                }

                if (destSubFolder != null)
                {
                    CopyFolders(sourceSubFolder, destSubFolder, ct, onProgress);
                    Marshal.ReleaseComObject(destSubFolder);
                }
                
                if (destFolders != null) Marshal.ReleaseComObject(destFolders);
                if (sourceSubFolder != null) Marshal.ReleaseComObject(sourceSubFolder);
            }
            if (sourceSubFolders != null) Marshal.ReleaseComObject(sourceSubFolders);
        }

        private Outlook.Folder FindFolderByName(Outlook.Folders folders, string name)
        {
            foreach (Outlook.Folder f in folders)
            {
                if (string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return f;
                }
                Marshal.ReleaseComObject(f);
            }
            return null;
        }

        private Outlook.Folder GetRootFolder(Outlook.NameSpace ns, string filePath, Action<int, string> onProgress)
        {
            // 1. Find the Store object first
            Outlook.Store targetStore = null;
            foreach (Outlook.Store store in ns.Stores)
            {
                if (string.Equals(store.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                {
                    targetStore = store;
                    break;
                }
            }

            if (targetStore != null)
            {
                // Try to get PR_IPM_SUBTREE_ENTRYID (0x35E00102)
                try
                {
                    const string PR_IPM_SUBTREE_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x35E00102";
                    object ipmProp = targetStore.PropertyAccessor.GetProperty(PR_IPM_SUBTREE_ENTRYID);
                    
                    string ipmEntryId = null;
                    if (ipmProp is string)
                    {
                        ipmEntryId = (string)ipmProp;
                    }
                    else if (ipmProp is byte[])
                    {
                        byte[] bytes = (byte[])ipmProp;
                        ipmEntryId = BitConverter.ToString(bytes).Replace("-", "");
                    }
                    
                    if (!string.IsNullOrEmpty(ipmEntryId))
                    {
                        var ipmRoot = ns.GetFolderFromID(ipmEntryId, targetStore.StoreID) as Outlook.Folder;
                        if (ipmRoot != null)
                        {
                            return ipmRoot;
                        }
                    }
                }
                catch (Exception ex)
                {
                     // Log warning only if verbose logging is enabled or critical
                     // onProgress(0, string.Format("Warning: Failed to resolve IPM Subtree: {0}. Implementation will fallback to Store Root.", ex.Message));
                }

                // Fallback to Store Root will happen in the legacy loop below
            }

            // Fallback: Legacy loop
            foreach (Outlook.Folder folder in ns.Folders)
            {
                try
                {
                    if (folder.Store != null)
                    {
                        if (string.Equals(folder.Store.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                            return folder;
                    }
                }
                catch { }
                Marshal.ReleaseComObject(folder);
            }

            return null;
        }
    }
}

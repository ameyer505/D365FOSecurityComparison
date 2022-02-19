using System;
using System.Collections.Generic;
using System.Text;

using System.IO.Compression;
using System.Security.Cryptography;


namespace D365FOSecurityComparison
{
    internal class Utility
    {
        internal static List<SecurityFile> getFiles(string path, ConsoleSpinner spinner = null)
        {
            List<SecurityFile> files = new List<SecurityFile>();

            using (ZipArchive Archive = ZipFile.OpenRead(path))
            {
                foreach (ZipArchiveEntry entry in Archive.Entries)
                {
                    LayerType type = LayerType.Role;
                    if (entry.FullName.ToLower().Contains("axsecurityrole"))
                        type = LayerType.Role;
                    if (entry.FullName.ToLower().Contains("axsecurityduty"))
                        type = LayerType.Duty;
                    if (entry.FullName.ToLower().Contains("axsecurityprivilege"))
                        type = LayerType.Privilege;

                    string hash = "";
                    using (var md5 = MD5.Create())
                    {
                        var hashByte = md5.ComputeHash(entry.Open());
                        hash = BitConverter.ToString(hashByte).Replace("-", "").ToLowerInvariant();
                    }

                    SecurityFile f = new SecurityFile()
                    {
                        Name = entry.Name,
                        Type = type,
                        Hash = hash
                    };

                    files.Add(f);

                    if (spinner != null) spinner.Turn();
                }
            }

            return files;
        }
    }
}

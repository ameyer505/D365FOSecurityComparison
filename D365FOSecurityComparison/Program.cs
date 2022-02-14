using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;

namespace D365FOSecurityComparison
{
    class Program
    {
        static void Main(string[] args)
        {
            List<SecurityFile> srcFiles = new List<SecurityFile>();
            List<SecurityFile> destFiles = new List<SecurityFile>();
            List<SecurityComparison> compFiles = new List<SecurityComparison>();

            if(args.Length != 2)
            {
                Console.WriteLine("Please provide security folder names in following syntax: <programName> <sourceFolder> <destFolder>");
            }
            else
            {
                using(ZipArchive srcArchive = ZipFile.OpenRead(args[0]))
                {
                    foreach(ZipArchiveEntry entry in srcArchive.Entries)
                    {
                        LayerType type = LayerType.Role;
                        if (entry.FullName.ToLower().Contains("axsecurityrole"))
                            type = LayerType.Role;
                        if (entry.FullName.ToLower().Contains("axsecurityduty"))
                            type = LayerType.Duty;
                        if (entry.FullName.ToLower().Contains("axsecurityprivilege"))
                            type = LayerType.Privilege;

                        string hash = "";
                        using(var md5 = MD5.Create())
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

                        srcFiles.Add(f);
                    }
                }

                using(ZipArchive destArchive = ZipFile.OpenRead(args[1]))
                {
                    foreach(ZipArchiveEntry entry in destArchive.Entries)
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

                        destFiles.Add(f);

                    }
                }
            }

            //Find security files added
            foreach (var destFile in destFiles)
            {
                if (!srcFiles.Any(sf => string.Equals(sf.Name, destFile.Name, StringComparison.CurrentCultureIgnoreCase) &&
                                          destFile.Type == sf.Type))
                {
                    SecurityComparison sc = new SecurityComparison()
                    {
                        Name = destFile.Name,
                        Type = destFile.Type,
                        Comparison = Action.Add
                    };

                    compFiles.Add(sc);
                }
            }

            //Find security files that have changed
            foreach(var destFile in destFiles)
            {
                SecurityFile comparisonFile = srcFiles.Where(sf => string.Equals(sf.Name, destFile.Name, StringComparison.CurrentCultureIgnoreCase) &&
                                          destFile.Type == sf.Type).FirstOrDefault();
                if(comparisonFile != null)
                {
                    if(!string.Equals(destFile.Hash, comparisonFile.Hash, StringComparison.CurrentCultureIgnoreCase))
                    {
                        SecurityComparison sc = new SecurityComparison()
                        {
                            Name = destFile.Name,
                            Type = destFile.Type,
                            Comparison = Action.Modify
                        };

                        compFiles.Add(sc);
                    }
                }
            }

            //Find security files removed
            foreach (var srcFile in srcFiles)
            {
               if(!destFiles.Any(df => string.Equals(df.Name, srcFile.Name, StringComparison.CurrentCultureIgnoreCase) &&
                                        srcFile.Type == df.Type))
                {
                    SecurityComparison sc = new SecurityComparison()
                    {
                        Name = srcFile.Name,
                        Type = srcFile.Type,
                        Comparison = Action.Remove
                    };

                    compFiles.Add(sc);
                }
            }

            foreach(var compFile in compFiles)
            {
                Console.WriteLine(compFile.Name + " - " + Enum.GetName(typeof(LayerType), compFile.Type) + " - " + Enum.GetName(typeof(Action), compFile.Comparison));
            }
            Console.Read();
        }
    }
}

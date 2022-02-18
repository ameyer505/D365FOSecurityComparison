using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

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
                Console.WriteLine("Processing source files");
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

                Console.WriteLine("Procesing destination files");
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

            Console.WriteLine("Comparing security files");
            Console.WriteLine("Processing added security");
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

            Console.WriteLine("Processing modified security");
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

            Console.WriteLine("Processing removed security");
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

            //foreach(var compFile in compFiles)
            //{
            //    Console.WriteLine(compFile.Name + " - " + Enum.GetName(typeof(LayerType), compFile.Type) + " - " + Enum.GetName(typeof(Action), compFile.Comparison));
            //}

            Console.WriteLine("Creating security document");
            //Creating the word document
            try
            {
                object styleHeading1 = "Heading 1";
                object styleHeading2 = "Heading 2";
                object missing = System.Reflection.Missing.Value;
                object endOfDoc = "\\endofdoc";

                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                word.ShowAnimation = false;
                word.Visible = false;
                
                Microsoft.Office.Interop.Word.Document doc = word.Documents.Add();

                //Roles Header 
                Microsoft.Office.Interop.Word.Paragraph paraRole = doc.Content.Paragraphs.Add(ref missing);
                paraRole.Range.set_Style(ref styleHeading1);
                paraRole.Range.Text = "Role";
                paraRole.Range.InsertParagraphAfter();

                //Roles Added
                var range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubRoleAdd = doc.Content.Paragraphs.Add(range);
                paraSubRoleAdd.Range.set_Style(ref styleHeading2);
                paraSubRoleAdd.Range.Text = "Added";
                paraSubRoleAdd.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtRoleAdded = doc.Content.Paragraphs.Add(range);
                StringBuilder rolesAdd = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Role && cf.Comparison == Action.Add))
                    rolesAdd.AppendLine(item.Name.Replace(".xml", ""));

                txtRoleAdded.Range.Text = rolesAdd.ToString();
                txtRoleAdded.Range.InsertParagraphAfter();

                //Roles Modified                    
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubRoleMod = doc.Content.Paragraphs.Add(range);
                paraSubRoleMod.Range.set_Style(ref styleHeading2);
                paraSubRoleMod.Range.Text = "Modified";
                paraSubRoleMod.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtRoleMod = doc.Content.Paragraphs.Add(range);
                StringBuilder rolesMod = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Role && cf.Comparison == Action.Modify))
                    rolesMod.AppendLine(item.Name.Replace(".xml", ""));

                txtRoleMod.Range.Text = rolesMod.ToString();
                txtRoleMod.Range.InsertParagraphAfter();

                //Roles Removed
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubRoleRem = doc.Content.Paragraphs.Add(range);
                paraSubRoleRem.Range.set_Style(ref styleHeading2);
                paraSubRoleRem.Range.Text = "Removed";
                paraSubRoleRem.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtRoleRem = doc.Content.Paragraphs.Add(range);
                StringBuilder rolesRem = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Role && cf.Comparison == Action.Remove))
                    rolesRem.AppendLine(item.Name.Replace(".xml", ""));

                txtRoleRem.Range.Text = rolesRem.ToString();
                txtRoleRem.Range.InsertParagraphAfter();

                //Duties Header
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph para2 = doc.Content.Paragraphs.Add(range);
                para2.Range.set_Style(ref styleHeading1);
                para2.Range.Text = "Duty";
                para2.Range.InsertParagraphAfter();

                //Duties Added
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubDutyAdd = doc.Content.Paragraphs.Add(range);
                paraSubDutyAdd.Range.set_Style(ref styleHeading2);
                paraSubDutyAdd.Range.Text = "Added";
                paraSubDutyAdd.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtDutyAdded = doc.Content.Paragraphs.Add(range);
                StringBuilder dutyAdd = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Duty && cf.Comparison == Action.Add))
                    dutyAdd.AppendLine(item.Name.Replace(".xml", ""));

                txtDutyAdded.Range.Text = rolesAdd.ToString();
                txtDutyAdded.Range.InsertParagraphAfter();

                //Duties Modified                    
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubDutyMod = doc.Content.Paragraphs.Add(range);
                paraSubDutyMod.Range.set_Style(ref styleHeading2);
                paraSubDutyMod.Range.Text = "Modified";
                paraSubDutyMod.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtDutyMod = doc.Content.Paragraphs.Add(range);
                StringBuilder dutiesMod = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Duty && cf.Comparison == Action.Modify))
                    dutiesMod.AppendLine(item.Name.Replace(".xml", ""));

                txtDutyMod.Range.Text = dutiesMod.ToString();
                txtDutyMod.Range.InsertParagraphAfter();

                //Duties Removed
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubDutyRem = doc.Content.Paragraphs.Add(range);
                paraSubDutyRem.Range.set_Style(ref styleHeading2);
                paraSubDutyRem.Range.Text = "Removed";
                paraSubDutyRem.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtDutyRem = doc.Content.Paragraphs.Add(range);
                StringBuilder dutyRem = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Duty && cf.Comparison == Action.Remove))
                    dutyRem.AppendLine(item.Name.Replace(".xml", ""));

                txtDutyRem.Range.Text = dutyRem.ToString();
                txtDutyRem.Range.InsertParagraphAfter();

                //Privilege Header  
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph para3 = doc.Content.Paragraphs.Add(range);
                para3.Range.set_Style(ref styleHeading1);
                para3.Range.Text = "Privilege";
                para3.Range.InsertParagraphAfter();

                //Privs Added
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubPrivAdd = doc.Content.Paragraphs.Add(range);
                paraSubPrivAdd.Range.set_Style(ref styleHeading2);
                paraSubPrivAdd.Range.Text = "Added";
                paraSubPrivAdd.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtPrivAdded = doc.Content.Paragraphs.Add(range);
                StringBuilder privAdd = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Privilege && cf.Comparison == Action.Add))
                    privAdd.AppendLine(item.Name.Replace(".xml", ""));

                txtPrivAdded.Range.Text = rolesAdd.ToString();
                txtPrivAdded.Range.InsertParagraphAfter();

                //Privs Modified
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubPrivMod = doc.Content.Paragraphs.Add(range);
                paraSubPrivMod.Range.set_Style(ref styleHeading2);
                paraSubPrivMod.Range.Text = "Modified";
                paraSubPrivMod.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtPrivMod = doc.Content.Paragraphs.Add(range);
                StringBuilder privsMod = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Privilege && cf.Comparison == Action.Modify))
                    privsMod.AppendLine(item.Name.Replace(".xml", ""));

                txtPrivMod.Range.Text = dutiesMod.ToString();
                txtPrivMod.Range.InsertParagraphAfter();

                //Privs Removed
                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Microsoft.Office.Interop.Word.Paragraph paraSubPrivRem = doc.Content.Paragraphs.Add(range);
                paraSubPrivRem.Range.set_Style(ref styleHeading2);
                paraSubPrivRem.Range.Text = "Removed";
                paraSubPrivRem.Range.InsertParagraphAfter();

                range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                var txtPrivRem = doc.Content.Paragraphs.Add(range);
                StringBuilder privRem = new StringBuilder();
                foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Privilege && cf.Comparison == Action.Remove))
                    dutyRem.AppendLine(item.Name.Replace(".xml", ""));

                txtPrivRem.Range.Text = privRem.ToString();
                txtPrivRem.Range.InsertParagraphAfter();

                //Save the doc
                string fileName = Directory.GetCurrentDirectory() + "\\"+args[0] + "_" + args[1]+"_Comparison.docx";
                doc.SaveAs2(fileName);
                doc.Close(missing, missing, missing);
                doc = null;
                word.Quit(missing, missing, missing);
                word = null;

                Console.WriteLine("Document created successfully!");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }

            Console.Read();
        }
    }
}

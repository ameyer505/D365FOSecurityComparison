using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            ConsoleSpinner spinner = new ConsoleSpinner();
            string fileType = "";

            string[] paths = new string[2];

            if (args.Length < 2 || args.Length > 3)
            {
                Console.WriteLine("Please provide security folder names in following syntax: <programName> <sourceFolder> <destFolder> <outputType>");
                return;
            }
            
            paths[0] = Path.GetFileNameWithoutExtension(args[0]);
            paths[1] = Path.GetFileNameWithoutExtension(args[1]);
            if(args.Length == 3)
            {
                fileType = args[2];
                if(!(string.Equals(fileType, "docx", StringComparison.CurrentCultureIgnoreCase) || string.Equals(fileType, "xlsx", StringComparison.CurrentCultureIgnoreCase)))
                {
                    Console.WriteLine("Unknown file type, please either user docx or xlsx options.");
                    return;
                }
            }
                

            Console.WriteLine("Processing source files");
            srcFiles = Utility.getFiles(args[0], spinner);

            Console.WriteLine("Procesing destination files");
            destFiles = Utility.getFiles(args[1], spinner);  

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
                    spinner.Turn();
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

                        spinner.Turn();
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

                    spinner.Turn();
                }
            }

            Console.WriteLine("Creating security document");
            Program p = new Program();
            //Create Word document
            if (fileType == "" || string.Equals(fileType, "docx", StringComparison.CurrentCultureIgnoreCase))
            {
                try
                {
                    p.CreateWordDocument(compFiles, paths[0], paths[1]);
                    Console.WriteLine("Document created successfully!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    throw ex;
                }
            }
            //Create Excel document
            else if(string.Equals(fileType, "xlsx", StringComparison.CurrentCultureIgnoreCase))
            {
                try
                {
                    p.CreateExcelDocument(compFiles, paths[0], paths[1]);
                    Console.WriteLine("Document created successfully!");
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    throw ex;
                }
            }
            else
            {
                Console.WriteLine("Unknown file type, please either use docx or xlsx options.");
                return;
            }


            Console.Read();
        }

        private bool CreateWordDocument(List<SecurityComparison> compFiles, string sourceFile, string destFile)
        {
            Microsoft.Office.Interop.Word.ParagraphFormat styleHeader = new Microsoft.Office.Interop.Word.ParagraphFormat();

            object missing = System.Reflection.Missing.Value;
            object endOfDoc = "\\endofdoc";

            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.ShowAnimation = false;
            word.Visible = false;

            Microsoft.Office.Interop.Word.Document doc = word.Documents.Add();

            //Add Document Header
            Microsoft.Office.Interop.Word.Paragraph paraHeader = doc.Content.Paragraphs.Add(ref missing);
            paraHeader.Range.Text = "Header";
            paraHeader.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1);
            paraHeader.Range.InsertParagraphAfter();

            var range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            var txtHeader = doc.Content.Paragraphs.Add(range);

            txtHeader.Range.Text = @"Create by program comparing security definitions between 2 versions of security source code. ";
            txtHeader.Range.InsertParagraphAfter();

            //Roles Header 
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Microsoft.Office.Interop.Word.Paragraph paraRole = doc.Content.Paragraphs.Add(range);
            paraRole.Range.Text = "Role";
            paraRole.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1);
            paraRole.Range.InsertParagraphAfter();

            //Roles Added
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Microsoft.Office.Interop.Word.Paragraph paraSubRoleAdd = doc.Content.Paragraphs.Add(range);
            paraSubRoleAdd.Range.Text = "Added";
            paraSubRoleAdd.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
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
            paraSubRoleMod.Range.Text = "Modified";
            paraSubRoleMod.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
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
            paraSubRoleRem.Range.Text = "Removed";
            paraSubRoleRem.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
            paraSubRoleRem.Range.InsertParagraphAfter();

            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            var txtRoleRem = doc.Content.Paragraphs.Add(range);
            StringBuilder rolesRem = new StringBuilder();
            foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Role && cf.Comparison == Action.Remove))
                rolesRem.AppendLine(item.Name.Replace(".xml", ""));

            if (compFiles.Where(cf => cf.Type == LayerType.Role && cf.Comparison == Action.Remove).Count() == 0)
                rolesRem.AppendLine("None");

            txtRoleRem.Range.Text = rolesRem.ToString();
            txtRoleRem.Range.InsertParagraphAfter();

            //Duties Header
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Microsoft.Office.Interop.Word.Paragraph para2 = doc.Content.Paragraphs.Add(range);
            para2.Range.Text = "Duty";
            para2.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1);
            para2.Range.InsertParagraphAfter();

            //Duties Added
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Microsoft.Office.Interop.Word.Paragraph paraSubDutyAdd = doc.Content.Paragraphs.Add(range);
            paraSubDutyAdd.Range.Text = "Added";
            paraSubDutyAdd.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
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
            paraSubDutyMod.Range.Text = "Modified";
            paraSubDutyMod.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
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
            paraSubDutyRem.Range.Text = "Removed";
            paraSubDutyRem.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
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
            para3.Range.Text = "Privilege";
            para3.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1);
            para3.Range.InsertParagraphAfter();

            //Privs Added
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Microsoft.Office.Interop.Word.Paragraph paraSubPrivAdd = doc.Content.Paragraphs.Add(range);
            paraSubPrivAdd.Range.Text = "Added";
            paraSubPrivAdd.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
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
            paraSubPrivMod.Range.Text = "Modified";
            paraSubPrivMod.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
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
            paraSubPrivRem.Range.Text = "Removed";
            paraSubPrivRem.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2);
            paraSubPrivRem.Range.InsertParagraphAfter();

            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            var txtPrivRem = doc.Content.Paragraphs.Add(range);
            StringBuilder privRem = new StringBuilder();
            foreach (var item in compFiles.Where(cf => cf.Type == LayerType.Privilege && cf.Comparison == Action.Remove))
                dutyRem.AppendLine(item.Name.Replace(".xml", ""));

            txtPrivRem.Range.Text = privRem.ToString();
            txtPrivRem.Range.InsertParagraphAfter();

            //Add Document Footer
            Microsoft.Office.Interop.Word.Paragraph paraFooter = doc.Content.Paragraphs.Add(ref missing);
            paraFooter.Range.Text = "Footer";
            paraFooter.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1);
            paraFooter.Range.InsertParagraphAfter();

            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            var txtFooter = doc.Content.Paragraphs.Add(range);

            txtHeader.Range.Text = @"Created by https://github.com/ameyer505/D365FOSecurityComparison.";
            txtHeader.Range.InsertParagraphAfter();

            //Save the doc
            string fileName = Directory.GetCurrentDirectory() + "\\" + sourceFile + "_" + destFile + "_Comparison.docx";
            doc.SaveAs2(fileName);
            doc.Close(missing, missing, missing);
            doc = null;
            word.Quit(missing, missing, missing);
            word = null;
            return true;
        }

        private bool CreateExcelDocument(List<SecurityComparison> compFiles, string sourceFile, string destFile)
        {
            using (var ep = new ExcelPackage())
            {
                //Roles
                var roleWS = ep.Workbook.Worksheets.Add("Roles");

                roleWS.Cells["A1"].Value = "Name";
                roleWS.Cells["A1"].Style.Font.Bold = true;
                roleWS.Cells["B1"].Value = "Action";
                roleWS.Cells["B1"].Style.Font.Bold = true;

                int i = 2;
                foreach(var addRole in compFiles.Where(sl => sl.Type == LayerType.Role && sl.Comparison == Action.Add).OrderBy(x => x.Name))
                {
                    roleWS.Cells["A" + i].Value = addRole.Name.Replace(".xml", "");
                    roleWS.Cells["B" + i].Value = "Add";
                    i++;
                }

                foreach (var modRole in compFiles.Where(sl => sl.Type == LayerType.Role && sl.Comparison == Action.Modify).OrderBy(x => x.Name))
                {
                    roleWS.Cells["A" + i].Value = modRole.Name.Replace(".xml", "");
                    roleWS.Cells["B" + i].Value = "Modify";
                    i++;
                }

                foreach (var remRole in compFiles.Where(sl => sl.Type == LayerType.Role && sl.Comparison == Action.Remove).OrderBy(x => x.Name))
                {
                    roleWS.Cells["A" + i].Value = remRole.Name.Replace(".xml", "");
                    roleWS.Cells["B" + i].Value = "Remove";
                    i++;
                }

                //Duties
                var dutyWS = ep.Workbook.Worksheets.Add("Duties");

                dutyWS.Cells["A1"].Value = "Name";
                dutyWS.Cells["A1"].Style.Font.Bold = true;
                dutyWS.Cells["B1"].Value = "Action";
                dutyWS.Cells["B1"].Style.Font.Bold = true;

                i = 2;
                foreach (var addDuty in compFiles.Where(sl => sl.Type == LayerType.Duty && sl.Comparison == Action.Add).OrderBy(x => x.Name))
                {
                    dutyWS.Cells["A" + i].Value = addDuty.Name.Replace(".xml", "");
                    dutyWS.Cells["B" + i].Value = "Add";
                    i++;
                }

                foreach (var modDuty in compFiles.Where(sl => sl.Type == LayerType.Duty && sl.Comparison == Action.Modify).OrderBy(x => x.Name))
                {
                    dutyWS.Cells["A" + i].Value = modDuty.Name.Replace(".xml", "");
                    dutyWS.Cells["B" + i].Value = "Modify";
                    i++;
                }

                foreach (var remDuty in compFiles.Where(sl => sl.Type == LayerType.Duty && sl.Comparison == Action.Remove).OrderBy(x => x.Name))
                {
                    dutyWS.Cells["A" + i].Value = remDuty.Name.Replace(".xml", "");
                    dutyWS.Cells["B" + i].Value = "Remove";
                    i++;
                }

                //Privileges
                var privWS = ep.Workbook.Worksheets.Add("Privileges");

                privWS.Cells["A1"].Value = "Name";
                privWS.Cells["A1"].Style.Font.Bold = true;
                privWS.Cells["B1"].Value = "Action";
                privWS.Cells["B1"].Style.Font.Bold = true;

                i = 2;
                foreach (var addPriv in compFiles.Where(sl => sl.Type == LayerType.Privilege && sl.Comparison == Action.Add).OrderBy(x => x.Name))
                {
                    privWS.Cells["A" + i].Value = addPriv.Name.Replace(".xml", "");
                    privWS.Cells["B" + i].Value = "Add";
                    i++;
                }

                foreach (var modPriv in compFiles.Where(sl => sl.Type == LayerType.Privilege && sl.Comparison == Action.Modify).OrderBy(x => x.Name))
                {
                    privWS.Cells["A" + i].Value = modPriv.Name.Replace(".xml", "");
                    privWS.Cells["B" + i].Value = "Modify";
                    i++;
                }

                foreach (var remPriv in compFiles.Where(sl => sl.Type == LayerType.Privilege && sl.Comparison == Action.Remove).OrderBy(x => x.Name))
                {
                    privWS.Cells["A" + i].Value = remPriv.Name.Replace(".xml", "");
                    privWS.Cells["B" + i].Value = "Remove";
                    i++;
                }

                //Save the file
                string fileName = Directory.GetCurrentDirectory() + "\\" + sourceFile + "_" + destFile + "_Comparison.xlsx";
                ep.SaveAs(new FileInfo(fileName));
            }

            return true;
        }
    }
}

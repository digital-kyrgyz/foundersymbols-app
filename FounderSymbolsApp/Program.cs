
    using System.Text;
    using System.Text.RegularExpressions;
    using FounderSymbolsApp.Dtos;
    using OfficeOpenXml;

    internal class Program
    {
        public static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    
            string projectPath = @"C:\Users\melis\source\projects\taxpayer-application";
            string excelFilePath = @"C:\Users\melis\Desktop\ExcelProject\Messages.xlsx";
    
            var messages = FindCyrillicMessages(projectPath);
            WriteToExcel(messages, excelFilePath);
    
            Console.WriteLine("Successfully searched and generated excel file");
            Console.ReadKey();

            static List<MessageInfo> FindCyrillicMessages(string path)
            {
                var messages = new List<MessageInfo>();
                foreach (string file in Directory.GetFiles(path, "*.*", SearchOption.AllDirectories).Where(f=>f.EndsWith(".cs") || f.EndsWith(".cshtml")))
                {
                    string[] lines = File.ReadAllLines(file);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (Regex.IsMatch(lines[i], @"\p{IsCyrillic}"))
                        {
                            messages.Add(new()
                            {
                                FilePath = file,
                                LineNumber = i + 1,
                                Message = lines[i]
                            });
                        }
                    }
                }

                return messages;
            }

            static void WriteToExcel(List<MessageInfo> messages, string filePath)
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Messages");

                    List<int> rowsToDelete = new List<int>();
            
                    worksheet.Cells[1, 1].Value = "Сообщение";
                    worksheet.Cells[1, 2].Value = "Путь к файлу";
                    worksheet.Cells[1, 3].Value = "Номер строки";

                    using (var range = worksheet.Cells[1, 1, 1, 3])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Font.Size = 14;
                    }
            
                    worksheet.Column(1).Width = 110;
                    worksheet.Column(2).Width = 75;
                    worksheet.Column(3).Width = 20;
            
                    for (int i = 0; i < messages.Count; i++)
                    {
                        string localProjectFolder = @"C:\Users\melis\source\projects\";
                        string relativePath = messages[i].FilePath.Replace(localProjectFolder, string.Empty);
                
                        string removedComment = RemoveLatinCharacters(messages[i].Message);
                        if (string.IsNullOrEmpty(removedComment))
                        {
                            rowsToDelete.Add(i + 2);
                            continue;;
                        }
                        worksheet.Cells[i + 2, 1].Value = removedComment;
                        worksheet.Cells[i + 2, 2].Value = relativePath;
                        worksheet.Cells[i + 2, 3].Value = messages[i].LineNumber;
                
                    }
            
                    for (int j = rowsToDelete.Count - 1; j >= 0; j--)
                    {
                        worksheet.DeleteRow(rowsToDelete[j]);
                    }
            
                    package.Save();
                }
            }
            static string RemoveLatinCharacters(string text)
            {
                string withoutComments = RemoveComments(text);

                StringBuilder builder = new StringBuilder();
                bool previousWasCyrillic = false;
        
                foreach (char c in withoutComments)
                {
                    if (IsCyrillic(c))
                    {
                        builder.Append(c);
                        previousWasCyrillic = true;
                    }
                    else
                    {
                        if (previousWasCyrillic)
                        {
                            builder.Append(' ');
                            previousWasCyrillic = false;
                        }
                    }
                }
        
                return builder.ToString();
            }

            static string RemoveComments(string text)
            {
                string pattern = @"//.*";
                return Regex.Replace(text, pattern, "");
            }
            static bool IsCyrillic(char c)
            {
                return (c >= 0x0400 && c <= 0x04FF) ||
                       (c >= 0x0500 && c <= 0x052F) || 
                       (c >= 0x2DE0 && c <= 0x2DFF) || 
                       (c >= 0xA640 && c <= 0xA69F);  
            }
        }
    }
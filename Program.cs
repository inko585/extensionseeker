using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExtensionSeeker
{
    public class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() != 2)
            {
                Console.WriteLine("Wrong amount of parameters, program call should look like this: ExtensionSeeker.exe [input folder] [output folder]");
            }

            var path1 = args[0];
            var path2 = args[1];

            var inputDict = new DirectoryInfo(path1);
            var outputDict = new DirectoryInfo(path2);

            if (!inputDict.Exists)
            {
                Console.WriteLine($"Could not find {path1}. Make sure the directory exists.");
            }

            if (!outputDict.Exists)
            {
                outputDict.Create();
            }

            var foundExt = 0;
            var missing = 0;

            var existing = outputDict.GetFiles().Select(x => x.Name);

            foreach (var fi in inputDict.GetFiles().Where(fi => fi.Directory != null))
            {
                var ext = GetFileExtension(File.ReadAllBytes(fi.FullName));


                if (ext != "")
                {
                    foundExt++;
                    if (!existing.Contains(Path.GetFileNameWithoutExtension(fi.Name) + ext))
                    {
                        fi.CopyTo(outputDict.FullName + "\\" + Path.GetFileNameWithoutExtension(fi.Name) + ext);                        
                    }
                } else
                {
                    missing++;
                }
            }

            Console.WriteLine($"Process completed: {foundExt} extension(s) found. {missing} file(s) were not allocatable.");

        }

        public static string GetFileExtension(IEnumerable<byte> bytes)
        {

            if (HasSignature(0, bytes, "4C-00-00-00-01-14-02-00"))
            {
                return ".LNK";
            }

            if (HasSignature(0, bytes, "37-7A-BC-AF-27-1C"))
            {
                return ".7z";
            }

            if (HasSignature(257, bytes, "75-73-74-61-72"))
            {
                return ".tar";
            }

            if (HasSignature(0, bytes, "25-50-44-46"))
            {
                return ".pdf";
            }

            if (HasSignature(0, bytes, "D0-CF-11-E0-A1-B1-1A-E1"))
            {

                if (HasSignature(512, bytes, "09-08-10-00-00-06-05-00"))
                {
                    return ".xls";
                }

                var workbookbyteseq = "00-00-57-00-6F-00-72-00-6B-00-62-00-6F-00-6F-00-6B-00-00";
                if (bytes.Count() > 10000 && (BitConverter.ToString(bytes.ToArray(), 0, 10000).Contains(workbookbyteseq) || BitConverter.ToString(bytes.ToArray(), bytes.Count() - 10000, 10000).Contains(workbookbyteseq)))
                {
                    return ".xls";
                }

                if (bytes.Count() < 10000 && BitConverter.ToString(bytes.ToArray(), 0, bytes.Count()).Contains(workbookbyteseq))
                {
                    return ".xls";
                }

                if (HasSignature(512, bytes, "EC-A5-C1-00"))
                {
                    return ".doc";
                }

                if (HasSignature(512, bytes, "00-6E-1E-F0"))
                {
                    return ".ppt";
                }

                if (HasSignature(512, bytes, "0F-00-E8-03"))
                {
                    return ".ppt";
                }
                if (HasSignature(512, bytes, "A0-46-1D-F03"))
                {
                    return ".ppt";
                }


                return ".msg";

            }

            if (HasSignature(0, bytes, "52-61-72-21-1A-07-00"))
            {
                return ".rar";
            }
            if (HasSignature(0, bytes, "52-61-72-21-1A-07-01-00"))
            {
                return ".rar";
            }
            if (HasSignature(0, bytes, "FF-D8"))
            {
                return ".jpg";
            }
            if (HasSignature(0, bytes, "89-50-4E-47-0D-0A-1A-0A"))
            {
                return ".png";
            }

            if (HasSignature(0, bytes, "3C-3F-78-6D-6C"))
            {
                return ".xml";
            }

            if (HasSignature(0, bytes, "50-4B-03-04"))
            {

                var rndFilePath = Path.GetTempPath() + RandomString(10) + ".zip";
                File.WriteAllBytes(rndFilePath, bytes.ToArray());

                var result = ".zip";
                try
                {
                    using (var zipFile = ZipFile.OpenRead(rndFilePath))
                    {

                        if (zipFile.Entries.Any(x => x.FullName.StartsWith("ppt/")))
                        {
                            if (zipFile.Entries.Any(x => Regex.IsMatch(x.FullName, @"ppt\/[^\/] *.bin")))
                            {
                                result = ".pptm";
                            }
                            else
                            {
                                result = ".pptx";
                            }
                        }
                        if (zipFile.Entries.Any(x => x.FullName.StartsWith("xl/")))
                        {
                            if (zipFile.Entries.Any(x => Regex.IsMatch(x.FullName, @"xl\/[^\/] *.bin")))
                            {
                                result = ".xlsm";
                            }
                            else
                            {
                                result = ".xlsx";
                            }
                        }
                        if (zipFile.Entries.Any(x => x.FullName.StartsWith("word/")))
                        {
                            if (zipFile.Entries.Any(x => Regex.IsMatch(x.FullName, @"word\/[^\/] *.bin")))
                            {
                                result = ".docm";
                            }
                            else
                            {
                                result = ".docx";
                            }
                        }
                    }
                    File.Delete(rndFilePath);
                }
                catch (Exception) { result = ".broken.zip"; }
                return result;
            }

            return "";

        }

        public static bool HasSignature(int offset, IEnumerable<byte> bytes, string byteSequence)
        {
            var sequenceSize = byteSequence.Split('-').Count();
            return BitConverter.ToString(bytes.Skip(offset).Take(sequenceSize).ToArray()) == byteSequence;
        }

        private static Random random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

    }
}

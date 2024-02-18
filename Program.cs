using Newtonsoft.Json;
using org.zefer.pd4ml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Razor.Tokenizer;
using static SupportClass;

namespace CoverLetter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello Pavan! are you generating Cover letter or email format press 1 for CL or 2 for EF?");
            int input = Convert.ToInt32(Console.ReadLine());

            if (input == 1)
            {
                string templatePath = "D:/MyProjects/Personal/CoverTemplate.html";
                string templateContent = System.IO.File.ReadAllText(templatePath);

                Console.WriteLine("Enter JSON data for Cover Letter:");
                string jsonData = Console.ReadLine();
                ClModel model = JsonConvert.DeserializeObject<ClModel>(jsonData);

                //{"City":"London","CompanyName":"ABC Corporation","Manager":"John Doe"}

                string replacedContent = templateContent
                    .Replace("[Current date]", DateTime.Now.ToLongDateString())
                    .Replace("[Hiring Manager]", model.Manager)
                    .Replace("[Company Name]", model.CompanyName)
                    .Replace("[Company City]", model.City);

                // Create an instance of Pd4Converter
                Pd4Converter pd4Converter = new Pd4Converter();

                // Convert the replaced content to PDF
                string pdfFileName = "CoverLetter.pdf";
                string pdfFilePath = pd4Converter.ConvertToPdf(pdfFileName, replacedContent, null, null);

                // Return the PDF for download
                Console.WriteLine($"Cover letter generated successfully. Press Enter to exit. PDF saved at: {pdfFilePath}");
                Console.ReadLine();

            }
            else if (input == 2)
            {
                // Email Format
                string emailTemplatePath = "D:/MyProjects/Personal/CoverLetter/EmailFormat.html";
                string emailTemplateContent = System.IO.File.ReadAllText(emailTemplatePath);

                // Take JSON input from the user
                Console.WriteLine("Enter JSON data for Email Format:");
                string jsonData = Console.ReadLine();
                EmailFormatDTO emailModel = JsonConvert.DeserializeObject<EmailFormatDTO>(jsonData);

                //{"CompanyName":"XYZ Tech Solutions","HRName":"Jane Smith","position":"Software Engineer"}

                string replacedContent = emailTemplateContent
                    .Replace("[position]", emailModel.position)
                    .Replace("[HRName]", emailModel.HRName)
                    .Replace("[CompanyName]", emailModel.CompanyName);

                // Create an instance of Pd4Converter
                Pd4Converter pd4Converter = new Pd4Converter();

                // Convert the replaced content to PDF
                string pdfFileName = "CoverLetter.pdf";
                string pdfFilePath = pd4Converter.ConvertToPdf(pdfFileName, replacedContent, null, null);

                // Return the PDF for download
                Console.WriteLine($"Cover letter generated successfully. Press Enter to exit. PDF saved at: {pdfFilePath}");
                Console.WriteLine("Press Enter to exit.");
                Console.ReadLine();
            }
        }
    }

    public class Pd4Converter
    {
        public string ConvertToPdf(string fileName, string html, string Password, string insets)
        {
            int width = 950;
            string orientation = "PORTRAIT";
            string pageSize = "A4";
            string filepath = "D://Files/" + fileName;

            while (File.Exists(filepath))//If file exists change file name
                filepath = filepath.Replace(".pdf", "_1.pdf");

            MemoryStream inputStream = new MemoryStream(ASCIIEncoding.Default.GetBytes(html));

            var outputStream = new System.IO.FileStream(filepath, System.IO.FileMode.Create);

            GeneratePdf(inputStream, width, pageSize, -1, null, orientation, insets, null, outputStream, Password);
            inputStream.Dispose();
            outputStream.Dispose();
            return filepath;
        }

        public void GeneratePdf(MemoryStream inputStream, int htmlWidth, String pageFormat, int permissions,
                                String bookmarks, String orientation, String insets, String fontsDir, Stream outfile,
                                 string password = null)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentUICulture = System.Threading.Thread.CurrentThread.CurrentCulture;

            PD4ML pd4Ml = new PD4ML();
            if (insets != null)
            {
                Tokenizer st = new Tokenizer(insets, ",");
                try
                {
                    int top = Int32.Parse(st.NextToken());
                    int left = Int32.Parse(st.NextToken());
                    int bottom = Int32.Parse(st.NextToken());
                    int right = Int32.Parse(st.NextToken());
                    String units = st.NextToken();
                    Rectangle ins = new Rectangle(top, left, bottom, right);
                    if ("MM".Equals(units.ToUpper()))
                    {
                        pd4Ml.PageInsetsMM = (ins);
                    }
                    else
                    {
                        pd4Ml.PageInsets = (ins);
                    }
                }
                catch (Exception e)
                {
                    e.ToString();
                    throw new Exception(
                            "Invalid page insets (top, left, bottom, right, units): "
                                    + insets);
                }
            }

            pd4Ml.HtmlWidth = (htmlWidth);
            if (!string.IsNullOrEmpty(password))
            {
                pd4Ml.setPermissions(password, 1, true);
            }

            if (orientation == null) orientation = "portrait";
            if ("PORTRAIT".Equals(orientation.ToUpper()))
                pd4Ml.PageSize = PD4Constants.getSizeByName(pageFormat);
            else
                pd4Ml.PageSize = pd4Ml.changePageOrientation(PD4Constants.getSizeByName(pageFormat));

            if (permissions != -1) pd4Ml.setPermissions("empty", permissions, true);
            if (bookmarks != null)
            {
                if ("ANCHORS".Equals(bookmarks.ToUpper()))
                    pd4Ml.generateOutlines(false);
                else if ("HEADINGS".Equals(bookmarks.ToUpper()))
                    pd4Ml.generateOutlines(true);
            }

            if (fontsDir != null && fontsDir.Length > 0) pd4Ml.useTTF(fontsDir, true);
            pd4Ml.render(inputStream, outfile);
        }
    }
}

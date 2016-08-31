using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Office.Interop.Word;

namespace OfficeToPdf.Web
{
    [ServiceContract]
    [ServiceBehavior(
        InstanceContextMode = InstanceContextMode.Single,
        AddressFilterMode = AddressFilterMode.Any,
        ConcurrencyMode = ConcurrencyMode.Single, IncludeExceptionDetailInFaults = true)]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    public class WebContract
    {
        [WebInvoke(Method = "POST", UriTemplate = "topdf", BodyStyle = WebMessageBodyStyle.Bare)]
        public Stream ConvertFrom(Stream input)
        {
            MultipartParser parser = new MultipartParser(input);

            if (parser.Success)
            {
                var randomGuid = System.Guid.NewGuid();
                var tempFilename = randomGuid + "-" + parser.Filename;
                using (
                    var stream =
                        new FileStream(Directory.GetCurrentDirectory() + @"\" + tempFilename, FileMode.Create, FileAccess.Write))
                {
                    stream.Write(parser.FileContents, 0, parser.FileContents.Length);
                }

                var appWord = new Microsoft.Office.Interop.Word.Application();
                var wordDocument = appWord.Documents.Open(Directory.GetCurrentDirectory() + @"\" + tempFilename);
                wordDocument.ExportAsFixedFormat(Directory.GetCurrentDirectory() + @"\" + tempFilename + ".pdf", WdExportFormat.wdExportFormatPDF);
                appWord.Documents.Close();

                var pdfBytes = File.ReadAllBytes(Directory.GetCurrentDirectory() + @"\" + tempFilename + ".pdf");

                try
                {
                    File.Delete(Directory.GetCurrentDirectory() + @"\" + tempFilename);
                    File.Delete(Directory.GetCurrentDirectory() + @"\" + tempFilename + ".pdf");
                } catch (Exception)
                { }

                WebOperationContext.Current.OutgoingResponse.Headers["Content-Disposition"] = "attachment; filename=\"" + parser.Filename + ".pdf\"";
                WebOperationContext.Current.OutgoingResponse.ContentType = "application/octet-stream";
                return new MemoryStream(pdfBytes);
            }
            else
            {
                throw new WebException(System.Net.HttpStatusCode.UnsupportedMediaType.ToString());
            }
        }
    }
}

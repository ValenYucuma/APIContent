using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace APIContent.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WordDocumentController : ControllerBase
    {
        private readonly IWebHostEnvironment _env;

        public WordDocumentController(IWebHostEnvironment env)
        {
            _env = env;
        }

        // 📌 1. Recibe el JSON y guarda el Word con encabezado y pie
        [HttpPost("insertar-header-footer")]
        public IActionResult InsertarHeaderFooter([FromBody] WordDocumentRequest request)
        {
            try
            {
                string folderPath = Path.Combine(_env.ContentRootPath, "ArchivosWord");
                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                string filePath = Path.Combine(folderPath, $"doc_{Guid.NewGuid()}.docx");

                // Guardar archivo original
                byte[] fileBytes = Convert.FromBase64String(request.ArchivoBase64);
                System.IO.File.WriteAllBytes(filePath, fileBytes);

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
                    if (mainPart.Document == null)
                        mainPart.Document = new Document(new Body());

                    // Crear e insertar header
                    HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
                    AddImageToHeader(headerPart, request.Encabezado);

                    // Crear e insertar footer
                    FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
                    AddImageToFooter(footerPart, request.PieDePagina);

                    // Asociar header y footer a la sección
                    SectionProperties sectionProps = mainPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
                    if (sectionProps == null)
                    {
                        sectionProps = new SectionProperties();
                        mainPart.Document.Body.Append(sectionProps);
                    }

                    sectionProps.RemoveAllChildren<HeaderReference>();
                    sectionProps.RemoveAllChildren<FooterReference>();

                    sectionProps.Append(
                        new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                        new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) }
                    );

                    mainPart.Document.Save();
                }

                return Ok(new { Ruta = filePath });
            }
            catch (Exception ex)
            {
                return BadRequest(new { Error = ex.Message, StackTrace = ex.StackTrace });
            }
        }

        // 📌 2. Devuelve el archivo desde la ruta
        [HttpPost("obtener-archivo")]
        public IActionResult ObtenerArchivo([FromBody] RutaRequest request)
        {
            if (!System.IO.File.Exists(request.Ruta))
                return NotFound(new { Error = "Archivo no encontrado" });

            byte[] fileBytes = System.IO.File.ReadAllBytes(request.Ruta);
            return File(fileBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                Path.GetFileName(request.Ruta));
        }

        // ----------------- Métodos auxiliares -----------------

        public static void AddImageToHeader(HeaderPart headerPart, string base64Image)
        {
            string rId = AddImagePart(headerPart, base64Image);

            Header header = new Header(
                new Paragraph(
                    new Run(
                        new Drawing(
                            CreateInlineImage(rId)
                        )
                    )
                )
            );
            headerPart.Header = header;
            headerPart.Header.Save();
        }

        public static void AddImageToFooter(FooterPart footerPart, string base64Image)
        {
            string rId = AddImagePart(footerPart, base64Image);

            Footer footer = new Footer(
                new Paragraph(
                    new Run(
                        new Drawing(
                            CreateInlineImage(rId)
                        )
                    )
                )
            );
            footerPart.Footer = footer;
            footerPart.Footer.Save();
        }

        private static string AddImagePart(OpenXmlPartContainer parentPart, string base64Image)
        {
            // ✅ Aquí usamos contentType en lugar de ImagePartType
            var imagePart = parentPart.AddNewPart<ImagePart>("image/png");

            byte[] imageBytes = Convert.FromBase64String(base64Image);
            using (var stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            return parentPart.GetIdOfPart(imagePart);
        }

        private static DW.Inline CreateInlineImage(string relationshipId)
        {
            long cx = 990000L; // ancho
            long cy = 792000L; // alto

            return new DW.Inline(
                new DW.Extent() { Cx = cx, Cy = cy },
                new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new D.GraphicFrameLocks() { NoChangeAspect = true }),
                new D.Graphic(
                    new D.GraphicData(
                        new Pic.Picture(
                            new Pic.NonVisualPictureProperties(
                                new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Image" },
                                new Pic.NonVisualPictureDrawingProperties()
                            ),
                            new Pic.BlipFill(
                                new D.Blip() { Embed = relationshipId },
                                new D.Stretch(new D.FillRectangle())
                            ),
                            new Pic.ShapeProperties(
                                new D.Transform2D(
                                    new D.Offset() { X = 0L, Y = 0L },
                                    new D.Extents() { Cx = cx, Cy = cy }),
                                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            );
        }
    }

    // DTOs
    public class WordDocumentRequest
    {
        public string ArchivoBase64 { get; set; } = "";
        public string Encabezado { get; set; } = "";  // Base64 imagen
        public string PieDePagina { get; set; } = ""; // Base64 imagen
    }

    public class RutaRequest
    {
        public string Ruta { get; set; } = "";
    }
}

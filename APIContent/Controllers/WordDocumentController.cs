using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;

// Namespaces para imágenes y dibujo
using A = DocumentFormat.OpenXml.Drawing;
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

        [HttpPost("insertar-header-footer")]
        public IActionResult InsertarHeaderFooter([FromBody] WordDocumentRequest request)
        {
            try
            {
                string folderPath = Path.Combine(_env.ContentRootPath, "ArchivosWord");
                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                string filePath = Path.Combine(folderPath, $"doc_{Guid.NewGuid()}.docx");

                // Guardar archivo base
                byte[] fileBytes = Convert.FromBase64String(request.ArchivoBase64);
                System.IO.File.WriteAllBytes(filePath, fileBytes);

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
                    if (mainPart.Document == null)
                        mainPart.Document = new Document(new Body());

                    // Crear header con tabla
                    HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
                    string headerImagePath = Path.Combine(_env.WebRootPath, "Plantillas", "header.jpeg");
                    string headerBase64 = Convert.ToBase64String(System.IO.File.ReadAllBytes(headerImagePath));
                    AddCustomHeader(headerPart, headerBase64, request);

                    // Crear footer
                    FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
                    string footerImagePath = Path.Combine(_env.WebRootPath, "Plantillas", "footer.jpeg");
                    string footerBase64 = Convert.ToBase64String(System.IO.File.ReadAllBytes(footerImagePath));
                    AddImageToFooter(footerPart, footerBase64);

                    // Asociar
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

        private static void AddCustomHeader(HeaderPart headerPart, string base64Image, WordDocumentRequest request)
        {
            // Agregamos la imagen (logo) desde el base64 que viene de InsertarHeaderFooter
            string rId = AddImagePart(headerPart, base64Image);

            Table table = new Table(
                new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                    ),
                    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
                ),

                // ==== FILA SUPERIOR ====
                new TableRow(
                    // Columna izquierda: títulos centrados + logo
                    new TableCell(                  
                        new Paragraph(new Run(new Drawing(CreateInlineImage(rId))))
                    ),
                    new TableCell(
                        new Paragraph(new Run(new Text($"Titulo: {request.TituloDocumento}")))
                    ),
                    // Columna derecha: Código, Versión, Página
                    new TableCell(
                        new Paragraph(new Run(new Text($"Código: {request.Codigo}"))),
                        new Paragraph(new Run(new Text($"Versión: {request.Version}"))),
                        new Paragraph(new Run(new Text($"Página: {request.Pagina}")))
                    )
                ),

                // ==== FILA INFERIOR (Firmas) ====
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text($"Elaboró: {request.Elaboro}\nFecha: {request.FechaElaboro}")))),
                    new TableCell(new Paragraph(new Run(new Text($"Revisó: {request.Reviso}\nFecha: {request.FechaReviso}")))),
                    new TableCell(new Paragraph(new Run(new Text($"Aprobó: {request.Aprobo}\nFecha: {request.FechaAprobo}"))))
                )
            );

            // Insertar la tabla en el header
            Header header = new Header(new Paragraph(new Run(table)));
            headerPart.Header = header;
            headerPart.Header.Save();
        }


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

        private static void AddImageToFooter(FooterPart footerPart, string base64Image)
        {
            string rId = AddImagePart(footerPart, base64Image);

            // Ajustar tamaño (ejemplo: 500px ancho, 100px alto)
            long widthEmu = 500 * 9525;  // 1 px = 9525 EMUs
            long heightEmu = 100 * 9525;

            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                        new DW.EffectExtent
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties { Id = (UInt32Value)1U, Name = "Footer Image" },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new Pic.Picture(
                                    new Pic.NonVisualPictureProperties(
                                        new Pic.NonVisualDrawingProperties
                                        {
                                            Id = (UInt32Value)0U,
                                            Name = "footer.jpeg"
                                        },
                                        new Pic.NonVisualPictureDrawingProperties()),
                                    new Pic.BlipFill(
                                        new A.Blip { Embed = rId },
                                        new A.Stretch(new A.FillRectangle())),
                                    new Pic.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset { X = 0L, Y = 0L },
                                            new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                                )
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });

            // Párrafo centrado
            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center }),
                new Run(element)
            );

            Footer footer = new Footer(paragraph);
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

    public class WordDocumentRequest
    {
        public string ArchivoBase64 { get; set; } = "";  // Documento base
        public string ImagenBase64 { get; set; } = "";   // Logo dinámico (ej: el escudo de la Gobernación)
        public string Codigo { get; set; } = "";
        public string Version { get; set; } = "";
        public string Pagina { get; set; } = "";
        public string Elaboro { get; set; } = "";
        public string Reviso { get; set; } = "";
        public string Aprobo { get; set; } = "";
        public string FechaElaboro { get; set; } = "";
        public string FechaReviso { get; set; } = "";
        public string FechaAprobo { get; set; } = "";
        public string TituloDocumento { get; set; } = "";
    }


    public class RutaRequest
    {
        public string Ruta { get; set; } = "";
    }
}

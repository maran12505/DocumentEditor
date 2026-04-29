namespace Deditor.Core.Models
{
    public class SaveParameter
    {
        public string? Content { get; set; }
        public string? FileName { get; set; }
    }

    public class CustomParameter
    {
        public string? Content { get; set; }
        public string? Type { get; set; }
    }

    public class RestrictParameter
    {
        public string? PasswordBase64 { get; set; }
        public string? SaltBase64 { get; set; }
        public int SpinCount { get; set; }
    }

    public class SpellCheckParameter
    {
        public string? Word { get; set; }
        public string? Language { get; set; }
    }

    public class MetafileParam
    {
        public string? action { get; set; }
        public string? imageData { get; set; }
        public string? format { get; set; }
    }

    /// <summary>
    /// DTO mirroring iPubEdit Meta Information fields, populated from a DOCX
    /// customXml part. Property names match IPubEditMeta in DocumentEditor.razor.
    /// </summary>
    public class IPubMetaDto
    {
        public string? XmlType { get; set; }
        public string? Publisher { get; set; }
        public string? JobCard { get; set; }
        public string? CeTemplate { get; set; }
        public string? BookOrJournalId { get; set; }
        public bool? Copyedit { get; set; }
        public bool? PreEditing { get; set; }
        public bool? Testing { get; set; }
        public string? Dtd { get; set; }
        public string? BookTitle { get; set; }
        public string? Isbn { get; set; }
        public string? Month { get; set; }
        public string? Year { get; set; }
        public string? PaginationPlatform { get; set; }
        public string? Country { get; set; }
        public string? ArticleType { get; set; }
        public string? Doi { get; set; }
        public string? Language { get; set; }
        public string? ProcessType { get; set; }
        public string? JournalName { get; set; }
        public string? JournalTitle { get; set; }
        public string? PIssn { get; set; }
        public string? EIssn { get; set; }
        public string? Copyrights { get; set; }
        public string? PublisherImprint { get; set; }
        public string? ArticleId { get; set; }
        public string? DocumentType { get; set; }

        // ── Article / chapter titles (XSD: <article-title>, <chapter-title>) ──
        public string? ArticleTitle { get; set; }
        public string? ChapterTitle { get; set; }

        // ── Issue-level numerics (XSD: <volume>, <issue>) ─────────────────────
        public string? Volume { get; set; }
        public string? Issue { get; set; }

        // ── Page range / e-location (XSD: <fpage>, <lpage>, <elocation-id>) ──
        public string? FirstPage { get; set; }
        public string? LastPage { get; set; }
        public string? ElocationId { get; set; }

        // ── Book series / edition (XSD: <edition>, <series>) ──────────────────
        public string? Edition { get; set; }
        public string? Series { get; set; }

        // ── Publication date day (XSD: <day> inside <pub-date>) ───────────────
        public string? Day { get; set; }

        // ── Publisher location (XSD: <publisher-loc>) — distinct from Country.
        // Was previously folded into Country; now captured separately so the
        // Country dropdown reflects only the country-of-publication field.
        public string? PublisherLoc { get; set; }
    }
}

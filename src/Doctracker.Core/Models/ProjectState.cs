using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace Doctracker.Core.Models
{
    [Serializable]
    [XmlRoot("DoctrackerProject")]
    public sealed class ProjectState
    {
        [XmlAttribute]
        public int SchemaVersion { get; set; } = 5;

        public string ProjectId { get; set; } = Guid.NewGuid().ToString("N");
        public List<CellLinkRecord> CellLinks { get; set; } = new List<CellLinkRecord>();
        public string SharedVaultPath { get; set; } = string.Empty;
        public string Revision { get; set; } = Guid.NewGuid().ToString("N");
        public List<XrefReservation> XrefReservations { get; set; } = new List<XrefReservation>();
        public string TestReference { get; set; } = string.Empty;
        public string WorkbookPath { get; set; } = string.Empty;
        public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;

        [XmlArrayItem("Document")]
        public List<DocumentRecord> Documents { get; set; } = new List<DocumentRecord>();

        [XmlArrayItem("Snip")]
        public List<SnipRecord> Snips { get; set; } = new List<SnipRecord>();

        [XmlArrayItem("Event")]
        public List<AuditEventRecord> AuditTrail { get; set; } = new List<AuditEventRecord>();
    }

    [Serializable]
    public sealed class CellLinkRecord
    {
        public string WorksheetName { get; set; } = string.Empty;
        public string CellAddress { get; set; } = string.Empty;
        public List<string> SnipIds { get; set; } = new List<string>();
    }

    [Serializable]
    public sealed class XrefReservation
    {
        public string Reference { get; set; } = string.Empty;
        public int Number { get; set; }
        public string DocumentId { get; set; } = string.Empty;
    }

    [Serializable]
    public sealed class DocumentRecord
    {
        [XmlAttribute]
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public List<string> Categories { get; set; } = new List<string>();
        public string TestReference { get; set; } = string.Empty;
        public int ReferenceNumber { get; set; }
        public long ByteLength { get; set; }
        public DateTime LastImportedAtUtc { get; set; }
        public string OriginalSourceName { get; set; } = string.Empty;
        public string OriginalSourceHash { get; set; } = string.Empty;
        public List<DocumentComment> Comments { get; set; } = new List<DocumentComment>();
        public string IndexKey { get; set; } = string.Empty;
        [XmlIgnore]
        public string DisplayName => string.IsNullOrWhiteSpace(TestReference) ? OriginalName : TestReference + " - " + ReferenceNumber.ToString("D2") + " - " + OriginalName;
        public string OriginalName { get; set; } = string.Empty;
        public string RelativePath { get; set; } = string.Empty;
        public string Sha256 { get; set; } = string.Empty;
        public int PageCount { get; set; }
        public bool IndexComplete { get; set; }
        public string IndexError { get; set; } = string.Empty;
        public DateTime AddedAtUtc { get; set; } = DateTime.UtcNow;

        private List<PageTextRecord> pages;
        [XmlIgnore]
        public Func<List<PageTextRecord>> PageLoader { get; set; }
        [XmlIgnore]
        public bool IndexDirty { get; private set; }
        [XmlArrayItem("Page")]
        public List<PageTextRecord> IndexedPages
        {
            get { return pages ?? (pages = PageLoader == null ? new List<PageTextRecord>() : PageLoader()); }
            set { pages = value; IndexDirty = true; }
        }
        public void MarkIndexSaved() { IndexDirty = false; }
        public void ReleaseIndex() { if (PageLoader != null && !IndexDirty) pages = null; }
    }

    [Serializable]
    public sealed class DocumentComment
    {
        // Font size in points on a reference page 595 points wide (A4).
        public double FontSize { get; set; } = 16;
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public int PageNumber { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string Text { get; set; } = string.Empty;
    }

    [Serializable]
    public sealed class PageTextRecord
    {
        [XmlAttribute]
        public int PageNumber { get; set; }

        // Preserve the V0.2 mixed-text representation when adding positional words.
        [XmlText]
        public string Text { get; set; } = string.Empty;

        [XmlArrayItem("Word")]
        public List<WordRecord> Words { get; set; } = new List<WordRecord>();
    }

    [Serializable]
    public sealed class WordRecord
    {
        public string Text { get; set; } = string.Empty;
        public int Line { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }

    [Serializable]
    public sealed class SnipRecord
    {
        [XmlAttribute]
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public string DocumentId { get; set; } = string.Empty;
        public int PageNumber { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public SnipType Type { get; set; }
        // Capture tool, independent from a table cell's parsed value type. Absent in older projects.
        public SnipType? SourceType { get; set; }
        public string RawText { get; set; } = string.Empty;
        public string ExtractedValue { get; set; } = string.Empty;
        public string WorksheetName { get; set; } = string.Empty;
        public string CellAddress { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public ReviewStatus Status { get; set; } = ReviewStatus.Prepared;
        public string PreparedBy { get; set; } = string.Empty;
        public DateTime PreparedAtUtc { get; set; } = DateTime.UtcNow;
        public string ReviewedBy { get; set; } = string.Empty;
        public DateTime? ReviewedAtUtc { get; set; }
    }

    [Serializable]
    public sealed class AuditEventRecord
    {
        [XmlAttribute]
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public DateTime AtUtc { get; set; } = DateTime.UtcNow;
        public string Actor { get; set; } = string.Empty;
        public string Action { get; set; } = string.Empty;
        public string EntityType { get; set; } = string.Empty;
        public string EntityId { get; set; } = string.Empty;
        public string Details { get; set; } = string.Empty;
    }

    public sealed class MatchCandidate
    {
        public string DocumentId { get; set; } = string.Empty;
        public int PageNumber { get; set; }
        public double Score { get; set; }
        public string Evidence { get; set; } = string.Empty;
        public bool IsPartial { get; set; }
        public bool IsExact { get; set; }
        public bool HasLocation { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; } = 1;
        public double Height { get; set; } = 1;
        public List<MatchCandidate> Fields { get; set; } = new List<MatchCandidate>();
    }
}

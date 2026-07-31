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
        public int SchemaVersion { get; set; } = 1;

        public string ProjectId { get; set; } = Guid.NewGuid().ToString("N");
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
    public sealed class DocumentRecord
    {
        [XmlAttribute]
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public string OriginalName { get; set; } = string.Empty;
        public string RelativePath { get; set; } = string.Empty;
        public string Sha256 { get; set; } = string.Empty;
        public int PageCount { get; set; }
        public DateTime AddedAtUtc { get; set; } = DateTime.UtcNow;

        [XmlArrayItem("Page")]
        public List<PageTextRecord> IndexedPages { get; set; } = new List<PageTextRecord>();
    }

    [Serializable]
    public sealed class PageTextRecord
    {
        [XmlAttribute]
        public int PageNumber { get; set; }

        [XmlText]
        public string Text { get; set; } = string.Empty;
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
    }
}

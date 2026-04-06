namespace Deditor.Core.Models
{
    public class ManuscriptRequest
    {
        public string? Sfdt { get; set; }
    }

    public class ManuscriptResponse
    {
        public string? Sfdt { get; set; }
        public int ChangesApplied { get; set; }
        public int TotalSuggestions { get; set; }
    }

    public class ChangeResponse
    {
        public List<ChangeItem>? Changes { get; set; }
    }

    public class ChangeItem
    {
        public string? OriginalText { get; set; }
        public string? ModifiedText { get; set; }
        public string? Reason { get; set; }
    }
}

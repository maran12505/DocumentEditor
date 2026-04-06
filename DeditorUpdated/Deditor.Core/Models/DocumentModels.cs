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
}

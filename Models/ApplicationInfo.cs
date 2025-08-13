namespace IntunePackagingTool
{
    public class IntuneApplication
    {
        public string DisplayName { get; set; } = "";
        public string Version { get; set; } = "";
        public string Category { get; set; } = "";
        public string Id { get; set; } = "";
        public string Publisher { get; set; } = "";
        public DateTime LastModified { get; set; }
    }

    public class ApplicationInfo
    {
        public string Manufacturer { get; set; } = "";
        public string Name { get; set; } = "";
        public string Version { get; set; } = "";
        public string InstallContext { get; set; } = "System";
        public string SourcesPath { get; set; } = "";
        public string ServiceNowSRI { get; set; } = "";
    }
}

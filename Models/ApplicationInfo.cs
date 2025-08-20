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
  public class IntuneWinInfo
    {
        public string FileName { get; set; } = "";
        public long UnencryptedContentSize { get; set; }
        public string EncryptedFilePath { get; set; } = "";
        public string TempDirectory { get; set; } = "";
        public EncryptionInfo EncryptionInfo { get; set; } = new();
    }

    public class EncryptionInfo
    {
        public string EncryptionKey { get; set; } = "";
        public string MacKey { get; set; } = "";
        public string InitializationVector { get; set; } = "";
        public string Mac { get; set; } = "";
        public string ProfileIdentifier { get; set; } = "";
        public string FileDigest { get; set; } = "";
        public string FileDigestAlgorithm { get; set; } = "";
    }

    public class AzureStorageInfo
    {
        public string SasUri { get; set; } = "";
    }
    
}

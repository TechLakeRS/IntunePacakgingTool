namespace IntunePackagingTool
{
    public enum DetectionRuleType
    {
        File,
        Registry,
        Script
    }

    public class DetectionRule
    {
        public DetectionRuleType Type { get; set; }
        public string Path { get; set; } = "";
        public string FileOrFolderName { get; set; } = "";
        public bool CheckVersion { get; set; }
        public string DetectionValue { get; set; } = "";
        public string Operator { get; set; } = "greaterThanOrEqual";
        
        // Registry specific properties
        public string RegistryHive { get; set; } = "";
        public string RegistryKey { get; set; } = "";
        public string RegistryValueName { get; set; } = "";
        public string ExpectedValue { get; set; } = "";
        
        // Script specific properties
        public string ScriptContent { get; set; } = "";
        public bool EnforceSignatureCheck { get; set; }
        public bool RunAs32Bit { get; set; }

        // Display properties for UI
        public string Icon => Type switch
        {
            DetectionRuleType.File => "📁",
            DetectionRuleType.Registry => "🗃️",
            DetectionRuleType.Script => "💻",
            _ => "❓"
        };

        public string Title => Type switch
        {
            DetectionRuleType.File => $"File Detection: {FileOrFolderName}",
            DetectionRuleType.Registry => $"Registry Detection: {RegistryValueName}",
            DetectionRuleType.Script => "PowerShell Script Detection",
            _ => "Unknown Detection"
        };

        public string Description => Type switch
        {
            DetectionRuleType.File => $"Path: {Path}\\{FileOrFolderName}" + (CheckVersion ? $" (Version {Operator} {DetectionValue})" : ""),
            DetectionRuleType.Registry => $"Key: {RegistryHive}\\{RegistryKey}\\{RegistryValueName}",
            DetectionRuleType.Script => "Custom PowerShell detection script",
            _ => ""
        };
    }
}
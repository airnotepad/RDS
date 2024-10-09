using System.Runtime.Versioning;

namespace RDS;

[SupportedOSPlatform("Windows")]
internal class Settings
{
    public string AnydeskPath { get; set; } = string.Empty;

    public void Save() => File.WriteAllText(Utils.SettingsPath, Newtonsoft.Json.JsonConvert.SerializeObject(this));
    public static Settings Get()
    {
        try
        {
            return Newtonsoft.Json.JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Utils.SettingsPath)) ?? new();
        }
        catch
        {
            return new();
        }
    }
}

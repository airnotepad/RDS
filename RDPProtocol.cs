using AxMSTSCLib;
using MSTSCLib;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;

namespace RDS;

public enum ERdpWindowResizeMode
{
    AutoResize = 0,
    Stretch = 1,
    Fixed = 2,
    StretchFullScreen = 3,
    FixedFullScreen = 4,
}

public enum ERdpFullScreenFlag
{
    Disable = 0,
    EnableFullScreen = 1,
    EnableFullAllScreens = 2,
}

public enum EDisplayPerformance
{
    /// <summary>
    /// Auto judge(by connection speed)
    /// </summary>
    Auto = 0,

    /// <summary>
    /// Low(8bit color with no feature support)
    /// </summary>
    Low = 1,

    /// <summary>
    /// Middle(16bit color with only font smoothing and desktop composition)
    /// </summary>
    Middle = 2,

    /// <summary>
    /// High(32bit color with full features support)
    /// </summary>
    High = 3,
}

public enum EGatewayMode
{
    AutomaticallyDetectGatewayServerSettings = 0,
    UseTheseGatewayServerSettings = 1,
    DoNotUseGateway = 2,
}

public enum EGatewayLogonMethod
{
    Password = 0,
    SmartCard = 1,
}

public enum EAudioRedirectionMode
{
    RedirectToLocal = 0,
    LeaveOnRemote = 1,
    Disabled = 2,
}

public enum EAudioQualityMode
{
    Dynamic = 0,
    Medium = 1,
    High = 2,
}

public class RdpLocalSetting
{
    public DateTime LastUpdateTime { get; set; } = DateTime.MinValue;
    public bool FullScreenLastSessionIsFullScreen { get; set; } = false;
    public int FullScreenLastSessionScreenIndex { get; set; } = -1;
}

public class RdpControlAdditionalSetting
{
    public string Name { get; set; } = "";
    public string? Value { get; set; } = "";
    public string ValueType { get; set; } = nameof(Int32);
    public string Description { get; set; } = "";
    public string HelpHrl { get; set; } = "";

    public T? GetValue<T>()
    {
        if (string.IsNullOrEmpty(Value))
            return default;
        return (T)Convert.ChangeType(Value, typeof(T));
    }
}

public sealed class RDP
{
    public string Address { get; set; } = "";
    public int Port { get; set; } = 0;
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public bool? IsAdministrativePurposes { get; set; } = false;
    public string Domain { get; set; } = "";
    public string LoadBalanceInfo { get; set; } = "";

    #region Display

    private ERdpFullScreenFlag? _rdpFullScreenFlag = ERdpFullScreenFlag.EnableFullScreen;
    public ERdpFullScreenFlag? RdpFullScreenFlag
    {
        get => _rdpFullScreenFlag;
        set
        {
            switch (value)
            {
                case ERdpFullScreenFlag.EnableFullAllScreens:
                    IsConnWithFullScreen = true;
                    if (RdpWindowResizeMode == ERdpWindowResizeMode.Fixed)
                        RdpWindowResizeMode = ERdpWindowResizeMode.FixedFullScreen;
                    if (RdpWindowResizeMode == ERdpWindowResizeMode.Stretch)
                        RdpWindowResizeMode = ERdpWindowResizeMode.StretchFullScreen;
                    break;

                case ERdpFullScreenFlag.Disable:
                    IsConnWithFullScreen = false;
                    if (RdpWindowResizeMode == ERdpWindowResizeMode.FixedFullScreen)
                        RdpWindowResizeMode = ERdpWindowResizeMode.Fixed;
                    if (RdpWindowResizeMode == ERdpWindowResizeMode.StretchFullScreen)
                        RdpWindowResizeMode = ERdpWindowResizeMode.Stretch;
                    break;

                case ERdpFullScreenFlag.EnableFullScreen:
                default:
                    break;
            }
        }
    }

    public bool? IsConnWithFullScreen { get; set; } = false;

    private bool? _isFullScreenWithConnectionBar = true;
    public bool? IsFullScreenWithConnectionBar
    {
        get => _isFullScreenWithConnectionBar;
        set
        {
            if (value == false)
            {
                IsPinTheConnectionBarByDefault = false;
            }
        }
    }

    private bool? IsPinTheConnectionBarByDefault { get; set; } = false;

    private ERdpWindowResizeMode? _rdpWindowResizeMode = ERdpWindowResizeMode.AutoResize;
    public ERdpWindowResizeMode? RdpWindowResizeMode
    {
        get => _rdpWindowResizeMode;
        set
        {
            var tmp = value;
            if (RdpFullScreenFlag == ERdpFullScreenFlag.Disable)
            {
                if (tmp == ERdpWindowResizeMode.FixedFullScreen)
                    tmp = ERdpWindowResizeMode.Fixed;
                if (tmp == ERdpWindowResizeMode.StretchFullScreen)
                    tmp = ERdpWindowResizeMode.Stretch;
            }
            _rdpWindowResizeMode = tmp;
        }
    }
    public int? RdpWidth { get; set; } = 800;
    public int? RdpHeight { get; set; } = 600;
    public bool? IsScaleFactorFollowSystem { get; set; } = true;

    private uint? _scaleFactorCustomValue = 100;
    public uint? ScaleFactorCustomValue
    {
        get => _scaleFactorCustomValue;
        set
        {
            uint? @new = value;
            if (value != null)
            {
                @new = (uint)value;
                if (@new > 300)
                    @new = 300;
                if (@new < 100)
                    @new = 100;
            }
            _scaleFactorCustomValue = @new;
        }
    }

    private EDisplayPerformance? DisplayPerformance { get; set; } = EDisplayPerformance.Auto;

    #endregion Display

    #region resource switch

    [DefaultValue(true)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnableClipboard { get; set; } = true;

    [DefaultValue(true)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnableDiskDrives { get; set; } = false;


    [DefaultValue(true)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnableRedirectDrivesPlugIn { get; set; } = false;


    [DefaultValue(true)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnableRedirectCameras { get; set; } = false;


    [DefaultValue(true)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnableKeyCombinations { get; set; } = true;


    [DefaultValue(EAudioRedirectionMode.RedirectToLocal)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public EAudioRedirectionMode? AudioRedirectionMode { get; set; } = EAudioRedirectionMode.RedirectToLocal;


    [DefaultValue(EAudioQualityMode.Dynamic)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public EAudioQualityMode? AudioQualityMode { get; set; } = EAudioQualityMode.Dynamic;


    [DefaultValue(false)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnableAudioCapture { get; set; } = false;


    [DefaultValue(false)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnablePorts { get; set; } = false;


    [DefaultValue(false)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnablePrinters { get; set; } = false;


    [DefaultValue(false)]
    [JsonProperty(DefaultValueHandling = DefaultValueHandling.Populate)]
    public bool? EnableSmartCardsAndWinHello { get; set; } = false;

    #endregion resource switch

    #region MSTSC model

    public bool MstscModeEnabled { get; set; } = false;

    public string RdpFileAdditionalSettings { get; set; } = "";

    #endregion

    #region RdpControlAdditionalSettings

    public string RdpControlAdditionalSettings { get; set; } = "";

    private static List<string>? _rdpControlAdditionalSettingKeys = null;
    public static List<string> GetRdpControlAdditionalSettingKeys()
    {
        if (_rdpControlAdditionalSettingKeys != null)
        {
            return _rdpControlAdditionalSettingKeys;
        }

        var excludeKeys = new HashSet<string>()
            {
                "Name", "Parent",
                nameof(AxMsRdpClient10.Server),
                nameof(AxMsRdpClient10.Domain),
                nameof(AxMsRdpClient10.UserName),
                nameof(IMsRdpClientAdvancedSettings8.RDPPort),
                nameof(AxMsRdpClient10.FullScreenTitle),
                nameof(AxMsRdpClient10.FullScreen),
                nameof(AxMsRdpClient10.DesktopHeight),
                nameof(AxMsRdpClient10.DesktopWidth),
                nameof(IMsRdpClientAdvancedSettings8.ConnectToAdministerServer),
                nameof(IMsRdpClientAdvancedSettings8.EnableMouse),
                nameof(IMsRdpClientAdvancedSettings8.LoadBalanceInfo),
            };


        // get all writable properties of AxMSTSCLib.AxMsRdpClient10/IMsRdpClientAdvancedSettings8 by reflection, which type is int or bool or string
        var keys = new List<string>();
        {
            {
                var type = typeof(IMsRdpClientAdvancedSettings8);
                var properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(p => p.CanWrite && (p.PropertyType == typeof(int) || p.PropertyType == typeof(bool) || p.PropertyType == typeof(string)));
                foreach (var propertyInfo in properties)
                {
                    if (excludeKeys.Contains(propertyInfo.Name)) continue;
                    string typeStr = ":s:";
                    if (propertyInfo.PropertyType == typeof(int))
                    {
                        typeStr = ":i:";
                    }
                    else if (propertyInfo.PropertyType == typeof(bool))
                    {
                        typeStr = ":i:";
                    }

                    keys.Add($"{propertyInfo.Name}{typeStr}");
                }
            }
            {
                var type = typeof(AxMSTSCLib.AxMsRdpClient10);
                var properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(p => p.CanWrite && (p.PropertyType == typeof(int) || p.PropertyType == typeof(bool) || p.PropertyType == typeof(string)));
                foreach (var propertyInfo in properties)
                {
                    if (excludeKeys.Contains(propertyInfo.Name)) continue;
                    string typeStr = ":s:";
                    if (propertyInfo.PropertyType == typeof(int))
                    {
                        typeStr = ":i:";
                    }
                    else if (propertyInfo.PropertyType == typeof(bool))
                    {
                        typeStr = ":i:";
                    }

                    keys.Add($"{propertyInfo.Name}{typeStr}");
                }
            }
        }
        _rdpControlAdditionalSettingKeys = keys.Distinct().OrderBy(x => x.ToLower()[0]).ToList();
        return _rdpControlAdditionalSettingKeys;
    }

    /// <summary>
    /// separate the rdpControlAdditionalSettings into `key`,`value`,`error message`, and `original string` tuples
    /// </summary>
    private static List<Tuple<string, string, string>> SplitAdditionalSettings(string rdpControlAdditionalSettings)
    {
        var results = new List<Tuple<string, string, string>>(); // return key, value, error message
        if (string.IsNullOrWhiteSpace(rdpControlAdditionalSettings) != false) return results;
        var separators = new[] { ":s:", ":i:", ":b:" };
        foreach (var s in rdpControlAdditionalSettings.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            int count = separators.Count(separator => s.IndexOf(separator, StringComparison.OrdinalIgnoreCase) >= 0);
            if (count != 1)
            {
                results.Add(new Tuple<string, string, string>(s, "", $"{s}: format error"));
            }
            else
            {
                foreach (var separator in separators)
                {
                    if (s.IndexOf(separator, StringComparison.OrdinalIgnoreCase) <= 0) continue;
                    var ss = s.Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
                    if (ss.Length != 2)
                    {
                        results.Add(new Tuple<string, string, string>(ss[0].Trim(), "", $"{s}: format error"));
                    }
                    else
                    {
                        var key = ss[0].Trim();
                        if (results.Any(x => x.Item1 == key))
                        {
                            results.Add(new Tuple<string, string, string>(key, "", $"{key}: duplicate key"));
                            break;
                        }
                        var value = ss[1].Trim();
                        switch (separator)
                        {
                            case ":i:":
                                results.Add(new Tuple<string, string, string>(key, value, int.TryParse(value, out var i) ? "" : $"{key}: value is not int"));
                                break;
                            case ":s:":
                                results.Add(new Tuple<string, string, string>(key, value, ""));
                                break;
                            case ":b:":
                                results.Add(new Tuple<string, string, string>(key, value, ""));
                                break;
                            default:
                                results.Add(new Tuple<string, string, string>(key, value, $"{key}: `{separator}` is not supported"));
                                break;
                        }
                    }
                    break;
                }
            }
        }
        return results;
    }

    public void ApplyRdpControlAdditionalSettings(AxMSTSCLib.AxMsRdpClient9NotSafeForScripting _rdpClient)
    {
        var sss = SplitAdditionalSettings(RdpControlAdditionalSettings);
        var propertiesAxMsRdpClient10 = typeof(AxMSTSCLib.AxMsRdpClient10).GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(p => p.CanWrite && (p.PropertyType == typeof(int) || p.PropertyType == typeof(bool) || p.PropertyType == typeof(string))).ToArray();
        var propertiesIMsRdpClientAdvancedSettings8 = typeof(IMsRdpClientAdvancedSettings8).GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(p => p.CanWrite && (p.PropertyType == typeof(int) || p.PropertyType == typeof(bool) || p.PropertyType == typeof(string))).ToArray();
        foreach (var tuple in sss)
        {
            if (tuple.Item3 != "") continue;
            if (GetRdpControlAdditionalSettingKeys().Any(x => x.StartsWith(tuple.Item1 + ":")) == false) continue;
            var key = tuple.Item1;
            var value = tuple.Item2;

            // AxMsRdpClient10
            {
                var pp = propertiesAxMsRdpClient10.FirstOrDefault(x => x.Name == key);
                if (pp != null && (pp.CanWrite || pp.SetMethod != null))
                {
                    if (pp.PropertyType == typeof(int))
                    {
                        if (int.TryParse(value, out var i))
                        {
                            pp.SetValue(_rdpClient, i);
                        }
                    }
                    else if (pp.PropertyType == typeof(bool))
                    {
                        if (int.TryParse(value, out var i))
                        {
                            pp.SetValue(_rdpClient, i > 0);
                        }
                    }
                    else if (pp.PropertyType == typeof(string))
                    {
                        pp.SetValue(_rdpClient, value);
                    }
                }
            }
            // IMsRdpClientAdvancedSettings8
            {
                var pp = propertiesIMsRdpClientAdvancedSettings8.FirstOrDefault(x => x.Name == key);
                if (pp != null && (pp.CanWrite || pp.SetMethod != null))
                {
                    if (pp.PropertyType == typeof(int))
                    {
                        if (int.TryParse(value, out var i))
                        {
                            pp.SetValue(_rdpClient.AdvancedSettings, i);
                        }
                    }
                    else if (pp.PropertyType == typeof(bool))
                    {
                        if (int.TryParse(value, out var i))
                        {
                            pp.SetValue(_rdpClient.AdvancedSettings, i > 0);
                        }
                    }
                    else if (pp.PropertyType == typeof(string))
                    {
                        pp.SetValue(_rdpClient.AdvancedSettings, value);
                    }
                }
            }
        }
    }

    #endregion

    #region Gateway
    public EGatewayMode? GatewayMode { get; set; } = EGatewayMode.DoNotUseGateway;
    public bool? GatewayBypassForLocalAddress { get; set; } = true;    
    public string GatewayHostName { get; set; } = "";
    public EGatewayLogonMethod? GatewayLogonMethod { get; set; } = EGatewayLogonMethod.Password;
    public string GatewayUserName { get; set; } = "";
    public string GatewayPassword { get; set; } = "";
    #endregion Gateway

    public RdpConfig ToRdpConfig()
    {
        var rdpConfig = new RdpConfig($"{this.Address}:{this.Port}", Username, Password, RdpFileAdditionalSettings)
        {
            Domain = this.Domain,
            LoadBalanceInfo = this.LoadBalanceInfo,
            AuthenticationLevel = 0,
            DisplayConnectionBar = this.IsFullScreenWithConnectionBar == true ? 1 : 0
        };

        switch (this.RdpFullScreenFlag)
        {
            case ERdpFullScreenFlag.Disable:
                rdpConfig.ScreenModeId = 1;
                rdpConfig.DesktopWidth = this.RdpWidth > 0 ? this.RdpWidth ?? 800 : 800;
                rdpConfig.DesktopHeight = this.RdpHeight > 0 ? this.RdpHeight ?? 600 : 600;
                break;

            case ERdpFullScreenFlag.EnableFullAllScreens:
                rdpConfig.ScreenModeId = 2;
                rdpConfig.UseMultimon = 1;
                break;

            case ERdpFullScreenFlag.EnableFullScreen:
                rdpConfig.ScreenModeId = 2;
                break;

            default:
                break;
        }

        switch (this.RdpWindowResizeMode)
        {
            case ERdpWindowResizeMode.Stretch:
                rdpConfig.SmartSizing = 1;
                rdpConfig.DynamicResolution = 0;
                break;

            case ERdpWindowResizeMode.Fixed:
                rdpConfig.SmartSizing = 0;
                rdpConfig.DynamicResolution = 0;
                rdpConfig.DesktopWidth = this.RdpWidth > 0 ? this.RdpWidth ?? 800 : 800;
                rdpConfig.DesktopHeight = this.RdpHeight > 0 ? this.RdpHeight ?? 600 : 600;
                break;

            case ERdpWindowResizeMode.AutoResize:
            default:
                rdpConfig.SmartSizing = 0;
                rdpConfig.DynamicResolution = 1;
                break;
        }

        rdpConfig.NetworkAutodetect = 0;
        switch (this.DisplayPerformance)
        {
            case EDisplayPerformance.Low:
                rdpConfig.ConnectionType = 1;
                rdpConfig.SessionBpp = 8;
                rdpConfig.AllowDesktopComposition = 0;
                rdpConfig.AllowFontSmoothing = 0;
                rdpConfig.DisableFullWindowDrag = 1;
                rdpConfig.DisableThemes = 1;
                rdpConfig.DisableWallpaper = 1;
                rdpConfig.DisableMenuAnims = 1;
                rdpConfig.DisableCursorSetting = 1;
                break;

            case EDisplayPerformance.Middle:
                rdpConfig.SessionBpp = 16;
                rdpConfig.ConnectionType = 3;
                rdpConfig.AllowDesktopComposition = 1;
                rdpConfig.AllowFontSmoothing = 1;
                rdpConfig.DisableFullWindowDrag = 1;
                rdpConfig.DisableThemes = 1;
                rdpConfig.DisableWallpaper = 1;
                rdpConfig.DisableMenuAnims = 1;
                rdpConfig.DisableCursorSetting = 1;
                break;

            case EDisplayPerformance.High:
                rdpConfig.SessionBpp = 32;
                rdpConfig.ConnectionType = 7;
                rdpConfig.AllowDesktopComposition = 1;
                rdpConfig.AllowFontSmoothing = 1;
                rdpConfig.DisableFullWindowDrag = 0;
                rdpConfig.DisableThemes = 0;
                rdpConfig.DisableWallpaper = 0;
                rdpConfig.DisableMenuAnims = 0;
                rdpConfig.DisableCursorSetting = 0;
                break;

            case EDisplayPerformance.Auto:
            default:
                rdpConfig.NetworkAutodetect = 1;
                break;
        }


        if (this.EnableDiskDrives == true)
        {
            rdpConfig.DriveStoreDirect = "*";
            rdpConfig.RedirectDrives = 1;
        }
        else
        {
            rdpConfig.DriveStoreDirect = "";
            rdpConfig.RedirectDrives = 0;
        }

        if (this.EnableRedirectDrivesPlugIn == true)
        {
            rdpConfig.RedirectDrives = 1;
            rdpConfig.DriveStoreDirect += ";DynamicDrives";
            rdpConfig.DriveStoreDirect = rdpConfig.DriveStoreDirect.Trim(';');
        }

        if (this.EnableClipboard == true)
            rdpConfig.RedirectClipboard = 1;
        if (this.EnablePrinters == true)
            rdpConfig.RedirectPrinters = 1;
        if (this.EnablePorts == true)
            rdpConfig.RedirectComPorts = 1;
        else
            rdpConfig.RedirectComPorts = 0;

        if (this.EnableSmartCardsAndWinHello == true)
            rdpConfig.RedirectSmartCards = 1;
        if (this.EnableKeyCombinations == true)
            rdpConfig.KeyboardHook = 2;
        else
            rdpConfig.KeyboardHook = 0;

        if (this.AudioRedirectionMode == EAudioRedirectionMode.RedirectToLocal)
            rdpConfig.AudioMode = 0;
        else if (this.AudioRedirectionMode == EAudioRedirectionMode.LeaveOnRemote)
            rdpConfig.AudioMode = 1;
        else if (this.AudioRedirectionMode == EAudioRedirectionMode.Disabled)
            rdpConfig.AudioMode = 2;

        if (this.AudioQualityMode == EAudioQualityMode.Dynamic)
            rdpConfig.AudioQualityMode = 0;
        else if (this.AudioQualityMode == EAudioQualityMode.Medium)
            rdpConfig.AudioQualityMode = 1;
        else if (this.AudioQualityMode == EAudioQualityMode.High)
            rdpConfig.AudioQualityMode = 2;

        if (this.EnableAudioCapture == true)
            rdpConfig.AudioCaptureMode = 1;

        rdpConfig.AutoReconnectionEnabled = 1;

        switch (GatewayMode)
        {
            case EGatewayMode.AutomaticallyDetectGatewayServerSettings:
                rdpConfig.GatewayUsageMethod = 2;
                break;

            case EGatewayMode.UseTheseGatewayServerSettings:
                rdpConfig.GatewayUsageMethod = 1;
                break;

            case EGatewayMode.DoNotUseGateway:
            default:
                rdpConfig.GatewayUsageMethod = 0;
                break;
        }
        rdpConfig.GatewayHostname = this.GatewayHostName;
        rdpConfig.GatewayCredentialsSource = 4;
        return rdpConfig;
    }
}

public sealed class RdpConfig
{
    /// <summary>
    /// (1) The remote session will appear in a window; (2) The remote session will appear full screen
    /// </summary>
    [RdpConfName("screen mode id:i:")]
    public int ScreenModeId { get; set; } = 2;

    [RdpConfName("use multimon:i:")]
    public int UseMultimon { get; set; } = 0;

    [RdpConfName("desktopwidth:i:")]
    public int DesktopWidth { get; set; } = 1600;

    [RdpConfName("desktopheight:i:")]
    public int DesktopHeight { get; set; } = 900;

    /// <summary>
    /// Determines whether or not the local computer scales the content of the remote session to fit the window size.
    /// 0 The local window content won't scale when resized
    /// 1 The local window content will scale when resized
    /// </summary>
    [RdpConfName("smart sizing:i:")]
    public int SmartSizing { get; set; } = 0;

    /// <summary>
    /// Determines whether the resolution of the remote session is automatically updated when the local window is resized.
    /// 0 Session resolution remains static for the duration of the session
    /// 1 Session resolution updates as the local window resizes
    /// </summary>
    [RdpConfName("dynamic resolution:i:")]
    public int DynamicResolution { get; set; } = 1;

    [RdpConfName("session bpp:i:")]
    public int SessionBpp { get; set; } = 32;

    /// <summary>
    /// winposstr:s:0,m,l,t,r,b
    /// m = mode ( 1 = use coords for window position, 3 = open as a maximized window )
    /// l = left
    /// t = top
    /// r = right  (ie Window width)
    /// b = bottom (ie Window height)
    /// winposstr:s:0,1,100,100,800,600  ---- Opens up a 800x600 window 100 pixels in from the left edge of your leftmost monitor and 100 pixels down from the upper edge.
    /// </summary>
    [RdpConfName("winposstr:s:")]
    public string Winposstr { get; set; } = "";

    /// <summary>
    /// Determines whether bulk compression is enabled when it is transmitted by RDP to the local computer.
    /// </summary>
    [RdpConfName("compression:i:")]
    public int Compression { get; set; } = 1;

    /// <summary>
    /// Determines when Windows key combinations (WIN key, ALT+TAB) are applied to the remote session for desktop connections.
    /// 0 Windows key combinations are applied on the local computer
    /// 1 Windows key combinations are applied on the remote computer when in focus
    /// 2 Windows key combinations are applied on the remote computer in full screen mode only
    /// </summary>
    [RdpConfName("keyboardhook:i:")]
    public int KeyboardHook { get; set; } = 2;

    /// <summary>
    /// Microphone redirection:Indicates whether audio input redirection is enabled.
    /// - 0: Disable audio capture from the local device
    /// - 1: Enable audio capture from the local device and redirection to an audio application in the remote session
    /// </summary>
    [RdpConfName("audiocapturemode:i:")]
    public int AudioCaptureMode { get; set; } = 0;

    /// <summary>
    /// Determines if the connection will use RDP-efficient multimedia streaming for video playback.
    /// - 0: Don't use RDP efficient multimedia streaming for video playback
    /// - 1: Use RDP-efficient multimedia streaming for video playback when possible
    /// </summary>
    [RdpConfName("videoplaybackmode:i:")]
    public int VideoPlaybackMode { get; set; } = 1;

    /// <summary>
    /// in old version, newer is networkautodetect
    /// The "connection tye" Remote Desktop option specifies which type of internet connection the remote connection is using, in terms of available bandwidth. Depending on the option you select, the Remote Desktop connection will change performance-related settings, including font smoothing, animations, Windows Aero, themes, desktop backgrounds, and so on.
    /// 1 Modem (56Kbps)
    /// 2 Low-speed broadband (256Kbps---2Mbps)
    /// 3 Satellite (2Mbps---16Mbps with high latency)
    /// 4 High-speed broadband (2Mbps---10Mbps)
    /// 5 WAN (10Mbps or higher with high latency)
    /// 6 LAN (10Mbps or higher)
    /// 7 Automatic bandwidth detection
    /// </summary>
    [RdpConfName("connection type:i:")]
    public int ConnectionType { get; set; } = 7;

    /// <summary>
    /// Determines whether automatic network type detection is enabled
    /// </summary>
    [RdpConfName("networkautodetect:i:")]
    public int NetworkAutodetect { get; set; } = 1;

    /// <summary>
    /// - 0: Disable automatic network type detection
    /// - 1: Enable automatic network type detection
    /// </summary>
    [RdpConfName("bandwidthautodetect:i:")]
    public int BandwidthAutodetect { get; set; } = 1;

    [RdpConfName("displayconnectionbar:i:")]
    public int DisplayConnectionBar { get; set; } = 1;

    [RdpConfName("disable wallpaper:i:")]
    public int DisableWallpaper { get; set; } = 1;

    [RdpConfName("allow font smoothing:i:")]
    public int AllowFontSmoothing { get; set; } = 0;

    [RdpConfName("allow desktop composition:i:")]
    public int AllowDesktopComposition { get; set; } = 0;

    [RdpConfName("disable full window drag:i:")]
    public int DisableFullWindowDrag { get; set; } = 1;

    [RdpConfName("disable menu anims:i:")]
    public int DisableMenuAnims { get; set; } = 1;

    [RdpConfName("disable themes:i:")]
    public int DisableThemes { get; set; } = 0;

    [RdpConfName("disable cursor setting:i:")]
    public int DisableCursorSetting { get; set; } = 0;

    /// <summary>
    /// This setting determines whether bitmaps are cached on the local computer. This setting corresponds to the selection in the Bitmap caching check box on the Experience tab of Remote Desktop Connection Options.
    /// </summary>
    [RdpConfName("bitmapcachepersistenable:i:")]
    public int BitmapCachePersistenable { get; set; } = 1;

    /// <summary>
    /// - 0: Play sounds on the local computer (Play on this computer)
    /// - 1: Play sounds on the remote computer(Play on remote computer)
    /// - 2: Do not play sounds(Do not play)
    /// </summary>
    [RdpConfName("audiomode:i:")]
    public int AudioMode { get; set; } = 1;

    /// <summary>
    /// - 0: Dynamic audio quality
    /// - 1: Medium audio quality
    /// - 2: High audio quality
    /// </summary>
    [RdpConfName("audioqualitymode:i:")]
    public int AudioQualityMode { get; set; } = 0;

    /// <summary>
    /// Determines whether the clipboard on the client computer will be redirected and available in the remote session and vice versa.
    /// 0 - Do not redirect the clipboard.
    /// 1 - Redirect the clipboard.
    /// </summary>
    [RdpConfName("redirectclipboard:i:")]
    public int RedirectClipboard { get; set; } = 0;

    [RdpConfName("redirectcomports:i:")]
    public int RedirectComPorts { get; set; } = 0;

    [RdpConfName("redirectdirectx:i:")]
    public int RedirectDirectX { get; set; } = 1;


    /// <summary>
    /// [2021-11-23 not work see #125]Determines whether local disk drives on the client computer will be redirected and available in the remote session.
    /// </summary>
    [RdpConfName("redirectdrives:i:")]
    public int RedirectDrives { get; set; } = 1;

    [RdpConfName("redirectposdevices:i:")]
    public int RedirectPosDevices { get; set; } = 0;

    [RdpConfName("redirectprinters:i:")]
    public int RedirectPrinters { get; set; } = 0;

    [RdpConfName("redirectsmartcards:i:")]
    public int RedirectSmartCards { get; set; } = 0;

    [RdpConfName("autoreconnection enabled:i:")]
    public int AutoReconnectionEnabled { get; set; } = 1;

    /// <summary>
    /// Defines the server authentication level settings.
    /// - 0: If server authentication fails, connect to the computer without warning (Connect and don't warn me)
    /// - 1: If server authentication fails, don't establish a connection (Don't connect)
    /// - 2: If server authentication fails, show a warning and allow me to connect or refuse the connection(Warn me)
    /// - 3: No authentication requirement specified.
    /// </summary>
    [RdpConfName("authentication level:i:")]
    public int AuthenticationLevel { get; set; } = 2;

    [RdpConfName("prompt for credentials:i:")]
    public int PromptForCredentials { get; set; } = 0;

    [RdpConfName("negotiate security layer:i:")]
    public int NegotiateSecurityLayer { get; set; } = 1;

    #region RemoteApp
    [RdpConfName("remoteapplicationmode:i:")]
    public int RemoteApplicationMode { get; set; } = 0;

    /// <summary>
    /// Specifies the name of the RemoteApp in the client interface while starting the RemoteApp.App display name. For example, "Excel 2016."
    /// </summary>
    [RdpConfName("remoteapplicationname:s:")]
    public string RemoteApplicationName { get; set; } = "";

    /// <summary>
    /// Specifies the alias or executable name of the RemoteApp. Valid alias or name. For example, "EXCEL."
    /// </summary>
    [RdpConfName("remoteapplicationprogram:s:")]
    public string RemoteApplicationProgram { get; set; } = "";

    /// <summary>
    /// Specifies whether the Remote Desktop client should check the remote computer for RemoteApp capabilities.
    /// 0 - Check the remote computer for RemoteApp capabilities before logging in.
    /// 1 - Do not check the remote computer for RemoteApp capabilities.Note: This setting must be set to 1 when connecting to Windows XP SP3, Vista or 7 computers with RemoteApps configured on them. This is the default behavior of RDP+.
    /// </summary>
    [RdpConfName("disableremoteappcapscheck:i:")]
    public int DisableRemoteAppCapsCheck { get; set; } = 1;

    [RdpConfName("alternate shell:s:")]
    public string AlternateShell { get; set; } = "";

    [RdpConfName("shell working directory:s:")]
    public string ShellWorkingDirectory { get; set; } = "";

    #endregion

    #region Gateway
    [RdpConfName("gatewayhostname:s:")]
    public string GatewayHostname { get; set; } = "";

    /// <summary>
    /// Specifies when to use an RD Gateway for the connection.
    /// - 0: Don't use an RD Gateway
    /// - 1: Always use an RD Gateway
    /// - 2: Use an RD Gateway if a direct connection cannot be made to the RD Session Host
    /// - 3: Use the default RD Gateway settings
    /// - 4: Don't use an RD Gateway, bypass gateway for local addresses
    /// Setting this property value to 0 or 4 are effectively equivalent, but setting this property to 4 enables the option to bypass local addresses.
    /// </summary>
    [RdpConfName("gatewayusagemethod:i:")]
    public int GatewayUsageMethod { get; set; } = 4;

    /// <summary>
    /// Specifies the RD Gateway authentication method.
    /// - 0: Ask for password (NTLM)
    /// - 1: Use smart card
    /// - 2: Use the credentials for the currently logged on user.
    /// - 3: Prompt the user for their credentials and use basic authentication
    /// - 4: Allow user to select later
    /// - 5: Use cookie-based authentication
    /// </summary>
    [RdpConfName("gatewaycredentialssource:i:")]
    public int GatewayCredentialsSource { get; set; } = 4;

    /// <summary>
    /// Specifies whether to use default RD Gateway settings.
    /// - 0: Use the default profile mode, as specified by the administrator
    /// - 1: Use explicit settings, as specified by the user
    /// </summary>
    [RdpConfName("gatewayprofileusagemethod:i:")]
    public int GatewayProfileUsageMethod { get; set; } = 0;
    #endregion

    [RdpConfName("promptcredentialonce:i:")]
    public int PromptCredentialOnce { get; set; } = 0;

    [RdpConfName("use redirection server name:i:")]
    public int UseRedirectionServerName { get; set; } = 0;

    [RdpConfName("rdgiskdcproxy:i:")]
    public int RdgiskdcProxy { get; set; } = 0;

    [RdpConfName("kdcproxyname:s:")]
    public string KdcProxyName { get; set; } = "";

    /// <summary>
    /// Determines which supported Plug and Play devices on the client computer will be redirected and available in the remote session.
    /// No value specified - Do not redirect any supported Plug and Play devices.
    /// * - Redirect all supported Plug and Play devices, including ones that are connected later.
    /// DynamicDevices - Redirect any supported Plug and Play devices that are connected later.
    /// The hardware ID for one or more Plug and Play devices - Redirect the specified supported Plug and Play device(s).
    /// </summary>
    [RdpConfName("devicestoredirect:s:")]
    public string DeviceStoreDirect { get; set; } = "*";

    /// <summary>
    /// Determines which local disk drives on the client computer will be redirected and available in the remote session.
    /// No value specified - Do not redirect any drives.
    /// * - Redirect all disk drives, including drives that are connected later.
    /// DynamicDrives - Redirect any drives that are connected later.
    /// The drive and labels for one or more drives - Redirect the specified drive(s). e.g. "C:\;D:\;"
    /// </summary>
    [RdpConfName("drivestoredirect:s:")]
    public string DriveStoreDirect { get; set; } = "*";

    /// <summary>
    /// Configures which cameras to redirect.
    /// This setting uses a semicolon-delimited list of KSCATEGORY_VIDEO_CAMERA interfaces of cameras enabled for redirection.
    /// "*" or ""
    /// </summary>
    [RdpConfName("camerastoredirect:s:")]
    public string CameraStoreDirect { get; set; } = "*";

    /// <summary>
    /// Specifies the name of the domain in which the user account that will be used to sign in to the remote computer is located.
    /// </summary>
    [RdpConfName("domain:s:")]
    public string Domain { get; set; } = "";

    /// <summary>
    /// loadbalanceinfo:s:tsv://MS Terminal Services Plugin.1.Wortell_sLab_Ses
    /// https://social.technet.microsoft.com/wiki/contents/articles/10392.rd-connection-broker-ha-and-the-rdp-properties-on-the-client.aspx
    /// </summary>
    [RdpConfName("loadbalanceinfo:s:")]
    public string LoadBalanceInfo { get; set; } = "";


    [RdpConfName("full address:s:")]
    public string FullAddress { get; set; } = "";

    [RdpConfName("username:s:")]
    public string Username { get; set; } = "";

    /// <summary>
    /// The user password in a binary hash value. Will be overruled by RDP+.
    /// </summary>
    [RdpConfName("password 51:b:")]
    public string Password { get; set; } = "";

    private readonly string _additionalSettings;

    public RdpConfig(string address, string username, string password, string additionalSettings = "")
    {
        FullAddress = address;
        Username = username;
        _additionalSettings = additionalSettings;

        if (string.IsNullOrEmpty(password) == false)
        {
            // encryption for rdp file
            Password = BitConverter.ToString(DataProtection.ProtectData(Encoding.Unicode.GetBytes(password), "")).Replace("-", "");
        }
    }

    public override string ToString()
    {
        var settings = new Dictionary<string, string>();

        // set all public properties by reflection
        foreach (var prop in typeof(RdpConfig).GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
        {
            foreach (RdpConfNameAttribute attr in prop.GetCustomAttributes(typeof(RdpConfNameAttribute), false))
            {
                var value = prop.GetValue(this)?.ToString();
                if (string.IsNullOrWhiteSpace(value) == false)
                    settings.Add(attr.Name, value);
            }
        }


        // set additional settings, if existed then replace
        if (string.IsNullOrWhiteSpace(_additionalSettings) == false)
        {
            foreach (var s in _additionalSettings.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                for (var i = 'a'; i <= 'z'; i++)
                {
                    if (Regex.IsMatch(s, @$":\s*{i}\s*:") == false) continue;
                    var ss = Regex.Split(s, $@":\s*{i}\s*:");
                    if (ss.Length == 2)
                    {
                        var key = $"{ss[0].Trim()}:{i}:";
                        var val = ss[1].Trim();
                        // if existed then replace
                        if (settings.ContainsKey(key))
                            settings[key] = val;
                        // or add
                        else
                            settings.Add(key, val);
                    }
                    break;
                }
            }
        }

        // if `selectedmonitors` is set, then force set the `screen mode id` to 2 and `use multimon:i:` to 1
        // 若设置了 `selectedmonitors`，则强制打开多显示器模式和全屏模式
        if (settings.ContainsKey("selectedmonitors:i:") && settings["selectedmonitors:i:"].Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Length > 1)
        {
            if (settings.ContainsKey("screen mode id:i:") == false)
                settings.Add("screen mode id:i:", "2");
            if (settings.ContainsKey("use multimon:i:") == false)
                settings.Add("use multimon:i:", "1");
            settings["screen mode id:i:"] = "2";
            settings["use multimon:i:"] = "1";
        }

        var str = new StringBuilder();
        foreach (var kv in settings)
        {
            str.AppendLine($"{kv.Key}{kv.Value}");
        }

        return str.ToString();
    }

    public static RdpConfig? FromRdpFile(string rdpFilePath)
    {
        var fi = new FileInfo(rdpFilePath);
        RdpConfig? rdpConfig = null;
        // read txt by line
        var pts = new[] { 's', 'i' };
        bool flag = false;
        if (fi.Exists)
        {
            rdpConfig = new RdpConfig(fi.Name.Replace(fi.Extension, ""), "", "", "");
            foreach (var line in System.IO.File.ReadLines(rdpFilePath))
            {
                //var ss = line.Split(":", StringSplitOptions.TrimEntries);
                foreach (var t in pts)
                {
                    if (Regex.IsMatch(line, @$":\s*{t}\s*:") == false) continue;
                    var ss = Regex.Split(line, $@":\s*{t}\s*:");
                    if (ss.Length == 2)
                    {
                        var key = $"{ss[0].Trim()}:{t}:";
                        var val = ss[1].Trim();
                        foreach (var prop in typeof(RdpConfig).GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).Where(x => x.Name != nameof(Password)))
                        {
                            if (prop.GetCustomAttributes(typeof(RdpConfNameAttribute), false).Cast<RdpConfNameAttribute>().Any(attr => string.Equals(attr.Name, key, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                flag = true;
                                if (t == 'i')
                                {
                                    if (int.TryParse(val, out var iVal))
                                    {
                                        prop.SetValue(rdpConfig, iVal);
                                    }
                                }
                                else
                                {
                                    prop.SetValue(rdpConfig, val);
                                }
                            }
                        }
                    }

                    break;
                }
            }
        }
        return flag ? rdpConfig : null;
    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    private class RdpConfNameAttribute : Attribute
    {
        public string Name { get; private set; }

        public RdpConfNameAttribute(string name)
        {
            this.Name = name;
        }
    }
}

[Serializable()]
internal sealed class DataProtection
{
    [Flags()]
    internal enum CryptProtectPromptFlags
    {
        CRYPTPROTECT_PROMPT_ON_UNPROTECT = 0x01,
        CRYPTPROTECT_PROMPT_ON_PROTECT = 0x02,
        CRYPTPROTECT_PROMPT_RESERVED = 0x04,
        CRYPTPROTECT_PROMPT_STRONG = 0x08,
        CRYPTPROTECT_PROMPT_REQUIRE_STRONG = 0x10
    }

    [Flags()]
    internal enum CryptProtectDataFlags
    {
        CRYPTPROTECT_UI_FORBIDDEN = 0x01,
        CRYPTPROTECT_LOCAL_MACHINE = 0x04,
        CRYPTPROTECT_CRED_SYNC = 0x08,
        CRYPTPROTECT_AUDIT = 0x10,
        CRYPTPROTECT_NO_RECOVERY = 0x20,
        CRYPTPROTECT_VERIFY_PROTECTION = 0x40,
        CRYPTPROTECT_CRED_REGENERATE = 0x80
    }

    internal static string ProtectData(string data, string name)
    {
        return ProtectData(data, name,
            CryptProtectDataFlags.CRYPTPROTECT_UI_FORBIDDEN | CryptProtectDataFlags.CRYPTPROTECT_LOCAL_MACHINE);
    }

    internal static byte[] ProtectData(byte[] data, string name)
    {
        return ProtectData(data, name,
            CryptProtectDataFlags.CRYPTPROTECT_UI_FORBIDDEN | CryptProtectDataFlags.CRYPTPROTECT_LOCAL_MACHINE);
    }

    internal static string ProtectData(string data, string name, CryptProtectDataFlags flags)
    {
        byte[] dataIn = Encoding.Unicode.GetBytes(data);
        byte[] dataOut = ProtectData(dataIn, name, flags);

        if (dataOut != null)
            return (Convert.ToBase64String(dataOut));
        else
            return null;
    }

    internal static byte[] ProtectData(byte[] data, string name, CryptProtectDataFlags dwFlags)
    {
        byte[] cipherText = null;

        // copy data into unmanaged memory
        DPAPI.DATA_BLOB din = new DPAPI.DATA_BLOB();
        din.cbData = data.Length;

        din.pbData = Marshal.AllocHGlobal(din.cbData);

        if (din.pbData.Equals(IntPtr.Zero))
            throw new OutOfMemoryException("Unable to allocate memory for buffer.");

        Marshal.Copy(data, 0, din.pbData, din.cbData);

        DPAPI.DATA_BLOB dout = new DPAPI.DATA_BLOB();

        try
        {
            bool cryptoRetval = DPAPI.CryptProtectData(ref din, name, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, dwFlags, ref dout);

            if (cryptoRetval)
            {
                int startIndex = 0;
                cipherText = new byte[dout.cbData];
                Marshal.Copy(dout.pbData, cipherText, startIndex, dout.cbData);
                DPAPI.LocalFree(dout.pbData);
            }
            else
            {
                int errCode = Marshal.GetLastWin32Error();
                StringBuilder buffer = new StringBuilder(256);
                Win32Error.FormatMessage(Win32Error.FormatMessageFlags.FORMAT_MESSAGE_FROM_SYSTEM, IntPtr.Zero, errCode, 0, buffer, buffer.Capacity, IntPtr.Zero);
            }
        }
        finally
        {
            if (!din.pbData.Equals(IntPtr.Zero))
                Marshal.FreeHGlobal(din.pbData);
        }

        return cipherText;
    }

    internal static void InitPromptstruct(ref DPAPI.CRYPTPROTECT_PROMPTSTRUCT ps)
    {
        ps.cbSize = Marshal.SizeOf(typeof(DPAPI.CRYPTPROTECT_PROMPTSTRUCT));
        ps.dwPromptFlags = 0;
        ps.hwndApp = IntPtr.Zero;
        ps.szPrompt = null;
    }
}

[SuppressUnmanagedCodeSecurity()]
internal class DPAPI
{
    [DllImport("crypt32")]
    public static extern bool CryptProtectData(ref DATA_BLOB dataIn, string szDataDescr, IntPtr optionalEntropy, IntPtr pvReserved,
        IntPtr pPromptStruct, DataProtection.CryptProtectDataFlags dwFlags, ref DATA_BLOB pDataOut);

    [DllImport("crypt32")]
    public static extern bool CryptUnprotectData(ref DATA_BLOB dataIn, StringBuilder ppszDataDescr, IntPtr optionalEntropy,
        IntPtr pvReserved, IntPtr pPromptStruct, DataProtection.CryptProtectDataFlags dwFlags, ref DATA_BLOB pDataOut);

    [DllImport("Kernel32.dll")]
    public static extern IntPtr LocalFree(IntPtr hMem);

    [StructLayout(LayoutKind.Sequential)]
    public struct DATA_BLOB
    {
        public int cbData;
        public IntPtr pbData;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct CRYPTPROTECT_PROMPTSTRUCT
    {
        public int cbSize; // = Marshal.SizeOf(typeof(CRYPTPROTECT_PROMPTSTRUCT))
        public int dwPromptFlags; // = 0
        public IntPtr hwndApp; // = IntPtr.Zero
        public string szPrompt; // = null
    }
}

internal class Win32Error
{
    [Flags()]
    internal enum FormatMessageFlags : int
    {
        FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x0100,
        FORMAT_MESSAGE_IGNORE_INSERTS = 0x0200,
        FORMAT_MESSAGE_FROM_STRING = 0x0400,
        FORMAT_MESSAGE_FROM_HMODULE = 0x0800,
        FORMAT_MESSAGE_FROM_SYSTEM = 0x1000,
        FORMAT_MESSAGE_ARGUMENT_ARRAY = 0x2000,
        FORMAT_MESSAGE_MAX_WIDTH_MASK = 0xFF,
    }

    [DllImport("Kernel32.dll")]
    internal static extern int FormatMessage(FormatMessageFlags flags, IntPtr source, int messageId, int languageId,
        StringBuilder buffer, int size, IntPtr arguments);
}

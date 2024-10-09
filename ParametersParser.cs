using System.Web;

namespace RDS;

internal enum Protocol
{
    RDP,
    Anydesk
}

internal class Patameters(string parameters)
{
    private string parameters = parameters;
    internal Protocol Protocol => GetProtocol();
    internal RDP RDP => GetRDP();
    internal Anydesk Anydesk => GetAnydesk();
    internal static Patameters Parse(IEnumerable<string> parameters)
    {
        var parameterString = parameters.First();
        parameterString = parameterString.Substring(parameterString.IndexOf("//") + 2).TrimEnd('/');

        return new(parameterString);
    }

    private Protocol GetProtocol()
    {
        var query = HttpUtility.ParseQueryString(HttpUtility.UrlDecode(parameters));

        var _protocol = query.Get("protocol") ?? throw new Exception("Protocol is empty");

        if (!Enum.TryParse<Protocol>(_protocol, out var protocol))
            throw new Exception("Protocol invalid value");

        return protocol;
    }

    private RDP GetRDP()
    {
        var query = HttpUtility.ParseQueryString(HttpUtility.UrlDecode(parameters));

        var address = query.Get("address") ?? throw new Exception("Adrress is empty");
        int.TryParse(query.Get("port") ?? throw new Exception("Port is empty"), out int port);
        var username = query.Get("username") ?? throw new Exception("Username is empty");
        var password = query.Get("password") ?? throw new Exception("Password is empty");

        if (port <= 0)
            throw new Exception("Post cannot be less zero");

        return new RDP()
        {
            Address = address,
            Port = port,
            Username = username,
            Password = password
        };
    }

    private Anydesk GetAnydesk()
    {
        var query = HttpUtility.ParseQueryString(HttpUtility.UrlDecode(parameters));

        var id = query.Get("id") ?? throw new Exception("Id is empty");
        var password = query.Get("password") ?? throw new Exception("Password is empty");

        return new Anydesk()
        {
            Id = id,
            Password = password,
        };
    }
}
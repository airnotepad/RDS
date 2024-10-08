using System.Web;

namespace RDS;

internal class Patameters(string parameters)
{
    private string parameters = parameters;
    internal RDP RDP => GetRDP();
    internal static Patameters Parse(IEnumerable<string> parameters)
    {
        var parameterString = parameters.First();
        parameterString = parameterString.Substring(parameterString.IndexOf("//") + 2).TrimEnd('/');

        return new(parameterString);
    }

    private RDP GetRDP()
    {
        var query = HttpUtility.ParseQueryString(HttpUtility.UrlDecode(parameters));

        var address = query.Get("address") ?? throw new Exception("Arress is empty");
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
}
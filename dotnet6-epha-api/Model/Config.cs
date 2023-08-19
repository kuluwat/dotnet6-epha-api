namespace  Model;

public static class Config
{
    private static IConfiguration StaticConfig { get; set; }

    public static void setConfig(IConfiguration IConfigurations)
    {
        StaticConfig = IConfigurations;
    }
    public static string Setting(string key)
    {
        return StaticConfig[key];
    }


}
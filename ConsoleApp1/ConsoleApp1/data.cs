using Newtonsoft.Json;

namespace ConsoleApp1;

public class Sql1
{
    [JsonProperty("jmeno")]
    public string Uzivatel { get; set; }
    [JsonProperty("taby")]
    public int Taby { get; set; }
}

public class Sql2
{
    public string NazevProcedury { get; set; }
    public string MappedValue { get; set; }
    public string NTUserName { get; set;}
    public int Count { get; set; }
}


public class Sql3
{
    public string NazevProcedury { get; set; }
    public string MappedValue { get; set; }
    public string NTUserName { get; set;}
    public string TimeOnly { get; set; }
}



using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Excel2Dataverse
{
  class JSONEntity
  {
    [Required] public string? Module { get; set; }
    [Required] public string? Name { get; set; }
    public Dictionary<string, string> Members { get; set; } = new();

    public string getSchemaName()
    {
      return $"{Module}.{Name}".ToLower();
    }

    public string getDisplayName()
    {
      return $"{Module}.{Name}";
    }

    public string getDisplayCollectionName()
    {
      return $"{Module}.{Name}s";
    }

    public EntityMetadata generateEntity(int locale, string description)
    {
      return new EntityMetadata
      {
        SchemaName = getSchemaName(),
        DisplayName = new Label(getDisplayName(), locale),
        DisplayCollectionName = new Label(getDisplayCollectionName(), locale),
        Description = new Label(description, locale),
        OwnershipType = OwnershipTypes.UserOwned,
        IsActivity = false
      };
    }
  }
}

class Program
{
  private static string connectionString()
  {
    ConnectionStringSettings settings = ConfigurationManager.ConnectionStrings["dynamics"];
    return settings.ConnectionString;
  }

  static void Main()
  {
    int defaultLocale = 2057;
    string defaultDescription = "Imported from the Ops App.";

    Console.WriteLine("Sign-in to your Microsoft Account ...\n\n");
    // IOrganizationService service = new ServiceClient(connectionString());
    // var whoAmI = (WhoAmIResponse)service.Execute(new WhoAmIRequest());
    // Console.WriteLine($"User ID is {whoAmI.UserId}.\n\n");

    Excel2Dataverse.JSONEntity testObject1 = new(){
      Module = "Countries",
      Name = "Country"
    };
    testObject1.Members.Add("Active", "Boolean");
    testObject1.Members.Add("Capital", "String");
    testObject1.Members.Add("Continent", "String");

    string testJSON1 = JsonSerializer.Serialize(testObject1);

    Console.WriteLine(testJSON1);

    Excel2Dataverse.JSONEntity? testObject2 = JsonSerializer.Deserialize<Excel2Dataverse.JSONEntity>(testJSON1);

    // Create entity
    // CreateEntityRequest createrequest = new CreateEntityRequest();

    // service.Execute(createrequest);
    // Console.WriteLine("The bank account entity has been created.");

    // Pause the console so it does not close.
    Console.WriteLine("Press the <Enter> key to exit.");
    Console.ReadLine();
  }
}

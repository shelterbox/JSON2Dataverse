using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using Excel2Dataverse;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Excel2Dataverse
{
  enum DataType
  {
    String = 0
  }

  class EntryPoint
  {
    public string? Prefix { get; set; }
    [Required] public List<Entity> Entities { get; set; } = new List<Entity>();
    public string? DefaultDescription { get; set; }
    // Locale ID
    // https://learn.microsoft.com/en-gb/previous-versions/windows/embedded/ms912047(v=winembedded.10)
    public int DefaultLocale { get; set; } = 2057;
  }

  class Entity
  {
    [JsonIgnore] public EntryPoint Parent { get; set; }
    [Required] public string? SchemaName { get; set; }
    [Required] public string? DisplayName { get; set; }
    [Required] public string? DisplayCollectionName { get; set; }
    public string? Description { get; set; }
    public Attribute? PrimaryAttribute { get; set; }
    public List<Attribute> Attributes { get; set; } = new List<Attribute>();

    public Entity(EntryPoint entryPoint)
    {
      entryPoint.Entities.Add(this);
      Parent = entryPoint;
    }

    public EntityMetadata generateEntity()
    {
      return new EntityMetadata
      {
        SchemaName = SchemaName,
        DisplayName = new Label(DisplayName, Parent.DefaultLocale),
        DisplayCollectionName = new Label(DisplayCollectionName, Parent.DefaultLocale),
        Description = new Label(Description, Parent.DefaultLocale),
        OwnershipType = OwnershipTypes.UserOwned,
        IsActivity = false
      };
    }
  }

  class Attribute
  {
    [JsonIgnore] public Entity Parent { get; set; }
    [Required] public string? SchemaName { get; set; }
    [Required] public string? DisplayName { get; set; }
    [Required] public DataType DataType { get; set; }
    public string? Description { get; set; }

    public Attribute(Entity entity)
    {
      entity.Attributes.Add(this);
      Parent = entity;
    }

    public AttributeMetadata generateAttribute()
    {
      switch (DataType)
      {
        case DataType.String:
          return new StringAttributeMetadata
          {
            SchemaName = SchemaName,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            MaxLength = 100,
            FormatName = StringFormatName.Text,
            DisplayName = new Label(DisplayName, Parent.Parent.DefaultLocale),
            Description = new Label(Description, Parent.Parent.DefaultLocale)
          };
        default:
          throw new Exception("Undefined 'DataType' property");
      }
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
    Console.WriteLine("Sign-in to your Microsoft Account ...\n\n");

    EntryPoint entry = new EntryPoint()
    {
      DefaultDescription = "Imported from the Ops app",
      Prefix = "OAP"
    };

    Excel2Dataverse.Entity entity = new Excel2Dataverse.Entity(entry)
    {
      DisplayName = "Events.Event",
      DisplayCollectionName = "Events.Events"
    };

    Excel2Dataverse.Attribute attribute = new Excel2Dataverse.Attribute(entity)
    {
      DataType = Excel2Dataverse.DataType.String,
      DisplayName = "ID",
      SchemaName = "ID"
    };

    entity.PrimaryAttribute = attribute;

    string theJson = JsonSerializer.Serialize(entry, new JsonSerializerOptions() { ReferenceHandler = ReferenceHandler.Preserve, IncludeFields = true });

    Console.WriteLine(theJson);

    EntryPoint? theEntry = JsonSerializer.Deserialize<EntryPoint>(theJson, new JsonSerializerOptions() { ReferenceHandler = ReferenceHandler.Preserve, IncludeFields = true });

    // ServiceClient implements IOrganizationService interface
    // IOrganizationService service = new ServiceClient(connectionString());

    // var whoAmI = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

    // Console.WriteLine($"User ID is {whoAmI.UserId}.\n\n");

    // XLWorkbook wb = new XLWorkbook("C:\\Users\\CarterMoorse\\Downloads\\Events_CM.xlsx");
    // IXLWorksheets ws = wb.Worksheets;
    // IXLWorksheet firstws = ws.First();
    // int rowCount = firstws.RowCount();

    // Create entity
    string prefix = "OAP";
    string sheetName = "Events.Event";
    string schemaName = $"{prefix}_{sheetName.ToLower().Replace(".", "")}";
    string displayName = $"{prefix}_{sheetName}";
    string displayCollectionName = $"{prefix}_{sheetName}s";
    string description = "Imported from the Operations App";

    CreateEntityRequest createrequest = new CreateEntityRequest();

    // service.Execute(createrequest);
    // Console.WriteLine("The bank account entity has been created.");

    // Pause the console so it does not close.
    Console.WriteLine("Press the <Enter> key to exit.");
    Console.ReadLine();
  }
}

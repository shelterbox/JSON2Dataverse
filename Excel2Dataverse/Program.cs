using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
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

    public static string formatPrefix(string? prefix)
    {
      return prefix != null && prefix.Trim() != "" ? $"{prefix}_" : "";
    }

    public string getSchemaName(string? prefix)
    {
      return $"{formatPrefix(prefix)}{Module}.{Name}".Replace(".", "").ToLower();
    }

    public string getDisplayName()
    {
      return $"{Module}.{Name}";
    }

    public string getDisplayCollectionName()
    {
      return $"{Module}.{Name}s";
    }

    public EntityMetadata generateEntity(int locale, string? prefix, string? description)
    {
      string displayName = getDisplayName();
      string DisplayCollectionName = getDisplayCollectionName();

      return new EntityMetadata
      {
        SchemaName = getSchemaName(prefix),
        LogicalName = displayName,
        LogicalCollectionName = DisplayCollectionName,
        DisplayName = new Label(displayName, locale),
        DisplayCollectionName = new Label(DisplayCollectionName, locale),
        Description = new Label(description, locale),
        OwnershipType = OwnershipTypes.UserOwned,
        IsActivity = false,
      };
    }

    public List<AttributeMetadata> generateAttributes(int locale, string? prefix)
    {
      List<AttributeMetadata> attributes = new();

      foreach (KeyValuePair<string, string> entry in Members)
      {
        if (entry.Key.ToLower() == "id")
          continue;
        
        string schemaName = $"{formatPrefix(prefix)}{entry.Key}".ToLower();
        string logicalName = $"{formatPrefix(prefix)}{entry.Key}";
        string displayName = entry.Key;

        switch (entry.Value.ToLower())
        {
          case "enum":
          case "string":
            attributes.Add(new StringAttributeMetadata()
            {
              SchemaName = schemaName,
              LogicalName = logicalName,
              DisplayName = new Label(displayName, locale),
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
              MaxLength = 200
            });
            break;
          case "boolean":
            attributes.Add(new BooleanAttributeMetadata()
            {
              SchemaName = schemaName,
              LogicalName = logicalName,
              DisplayName = new Label(displayName, locale),
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
              OptionSet = new BooleanOptionSetMetadata(
                new OptionMetadata(new Label("True", locale), 1),
                new OptionMetadata(new Label("False", locale), 0)
              )
            });
            break;
          case "datetime":
            attributes.Add(new DateTimeAttributeMetadata()
            {
              SchemaName = schemaName,
              LogicalName = logicalName,
              DisplayName = new Label(displayName, locale),
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
              Format = DateTimeFormat.DateAndTime,
              ImeMode = ImeMode.Auto
            });
            break;
          case "integer":
            attributes.Add(new IntegerAttributeMetadata()
            {
              SchemaName = schemaName,
              LogicalName = logicalName,
              DisplayName = new Label(displayName, locale),
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
              Format = IntegerFormat.None
            });
            break;
          case "long":
            attributes.Add(new BigIntAttributeMetadata()
            {
              SchemaName = schemaName,
              LogicalName = logicalName,
              DisplayName = new Label(displayName, locale),
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None)
            });
            break;
          case "decimal":
            attributes.Add(new DecimalAttributeMetadata()
            {
              SchemaName = schemaName,
              LogicalName = logicalName,
              DisplayName = new Label(displayName, locale),
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
              Precision = 10
            });
            break;
          case "binary":
            attributes.Add(new FileAttributeMetadata()
            {
              SchemaName = schemaName,
              LogicalName = logicalName,
              DisplayName = new Label(displayName, locale),
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None)
            });
            break;
          case "auto number":
            attributes.Add(new IntegerAttributeMetadata()
            {
              SchemaName = schemaName,
              LogicalName = logicalName,
              DisplayName = new Label(displayName, locale),
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
              AutoNumberFormat = ""
            });
            break;
          default:
            throw new Exception($"Module: {Module}, Name: {Name}, Member: {entry.Key} - Data type value '{entry.Value}' does not exist.");
        }
      }

      return attributes;
    }

    public StringAttributeMetadata generatePrimaryAttribute(int locale, string? prefix)
    {
      return new StringAttributeMetadata()
      {
        SchemaName = $"{formatPrefix(prefix)}ID".ToLower(),
        DisplayName = new Label("ID", locale),
        MaxLength = 100,
        FormatName = StringFormatName.Text,
        RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.SystemRequired)
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
    /* TODO -- Add paths to:
                - Create entities (w/ .JSON)
                - Create relationships (w/ .JSON)
                - Uploade data (w/ .XLSX)

    */

    // Default parameters
    int defaultLocale = 1033; // 1033 or 2057 - for English
    string defaultPrefix = "OAP";
    string defaultDescription = "Imported from the Ops App.";


    // Login to Microsoft
    Console.WriteLine("Sign-in to your Microsoft Account ...\n");
    IOrganizationService service = new ServiceClient(connectionString());
    var whoAmI = (WhoAmIResponse)service.Execute(new WhoAmIRequest());
    Console.WriteLine($"User ID is {whoAmI.UserId}.\n");


    // Load entities JSON
    List<Excel2Dataverse.JSONEntity>? jsonEntities = JsonSerializer.Deserialize<List<Excel2Dataverse.JSONEntity>>(@"[
	{
		""Module"": ""Countries"",
		""Name"": ""Country"",
		""Members"": {
			""Active"": ""Boolean"",
			""Capital"": ""String"",
			""Continent"": ""String"",
			""DS"": ""String"",
			""EDGAR"": ""String"",
			""FIFA"": ""String"",
			""FIPS"": ""String"",
			""GAUL"": ""String"",
			""Geoname_ID"": ""Decimal"",
			""hasCountryImage"": ""Boolean"",
			""HealthRiskSeverity"": ""Enum"",
			""ID"": ""Long"",
			""ImageURL"": ""String"",
			""Intermediate_Region_Code"": ""String"",
			""Intermediate_Region_Name"": ""String"",
			""IOC"": ""String"",
			""ISO3166_1_Alpha_2"": ""String"",
			""ISO3166_1_Alpha_3"": ""String"",
			""ISO3166_1_Numeric"": ""String"",
			""ITU"": ""String"",
			""Latitude"": ""Decimal"",
			""Longitude"": ""Decimal"",
			""M49"": ""Decimal"",
			""MARC"": ""String"",
			""Name"": ""String"",
			""Name_ar_Formal"": ""String"",
			""Name_ar_Official"": ""String"",
			""Name_ar_Short"": ""String"",
			""Name_cn_Formal"": ""String"",
			""Name_cn_Official"": ""String"",
			""Name_cn_Short"": ""String"",
			""Name_en_Formal"": ""String"",
			""Name_en_Official"": ""String"",
			""Name_en_Short"": ""String"",
			""Name_es_Formal"": ""String"",
			""Name_es_Official"": ""String"",
			""Name_es_Short"": ""String"",
			""Name_fr_Formal"": ""String"",
			""Name_fr_Official"": ""String"",
			""Name_fr_Short"": ""String"",
			""Name_ru_Formal"": ""String"",
			""Name_ru_Official"": ""String"",
			""Name_ru_Short"": ""String"",
			""Region_Code"": ""String"",
			""Region_Name"": ""String"",
			""Sub_region_Code"": ""String"",
			""Sub_region_Name"": ""String"",
			""TimeZones"": ""String"",
			""TLD"": ""String"",
			""WMO"": ""String""
		}
	},
	{
		""Module"": ""Countries"",
		""Name"": ""CountryImage"",
		""Members"": {
			""changedDate"": ""DateTime"",
			""Contents"": ""Binary"",
			""createdDate"": ""DateTime"",
			""DeleteAfterDownload"": ""Boolean"",
			""FileID"": ""Auto number"",
			""__FileName__"": ""Long"",
			""HasContents"": ""Boolean"",
			""ID"": ""Long"",
			""Name"": ""String"",
			""PublicThumbnailPath"": ""String"",
			""Size"": ""Long"",
			""__UUID__"": ""String""
		}
	}
]");

    if (jsonEntities == null)
      throw new Exception("No JSON to load - Invalid JSON.");

    Console.WriteLine($"{jsonEntities.Count} {(jsonEntities.Count == 1 ? "entity" : "entities")} loaded. Creating schemas ...");


    // Create entities
    int entitiyCount = 1;
    foreach (Excel2Dataverse.JSONEntity jsonEntity in jsonEntities)
    {
      EntityMetadata dataverseEntity = jsonEntity.generateEntity(defaultLocale, defaultPrefix, defaultDescription);
      StringAttributeMetadata dataversePrimaryAttribute = jsonEntity.generatePrimaryAttribute(defaultLocale, defaultPrefix);
      List<AttributeMetadata> dataverseAttributes = jsonEntity.generateAttributes(defaultLocale, defaultPrefix);

      CreateEntityRequest entityRequest = new()
      {
        Entity = dataverseEntity,
        HasActivities = false,
        HasFeedback = false,
        HasNotes = false,
        PrimaryAttribute = dataversePrimaryAttribute
      };

      service.Execute(entityRequest);
      Console.WriteLine($"{entitiyCount}/{jsonEntities.Count} - {dataverseEntity.LogicalName} entity schema has been created. Adding {dataverseAttributes.Count} attribute{(dataverseAttributes.Count == 1 ? "" : "s")}:");

      int attributeCount = 1;
      foreach (AttributeMetadata dataverseAttribute in dataverseAttributes)
      {
        CreateAttributeRequest attributeRequest = new()
        {
          Attribute = dataverseAttribute,
          EntityName = dataverseEntity.SchemaName
        };

        service.Execute(attributeRequest);
        Console.WriteLine($"\t{attributeCount}/{dataverseAttributes.Count} - {dataverseAttribute.LogicalName} attribute has been added.");
        attributeCount++;
      }

      entitiyCount++;
    }

    // TODO -- Create relationships

    // Pause the console so it does not close.
    Console.WriteLine("Press the <Enter> key to exit.");
    Console.ReadLine();
  }
}

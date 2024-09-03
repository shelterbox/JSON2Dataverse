using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Text.Json;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Excel2Dataverse
{
  class JSONEntity
  {
    [Required] public string? Name { get; set; }
    [Required] public string? PrimaryKey { get; set; }
    public Dictionary<string, string> Members { get; set; } = new();
    public Dictionary<string, string> ManyToOne { get; set; } = new();
    public Dictionary<string, string> ManyToMany { get; set; } = new();

    public static string addPrefix(string? prefix, string? text)
    {
      if (text == null)
        throw new Exception("'text' cannot be null");

      if (prefix == null || prefix.Trim() == "")
        return text;

      return $"{prefix}_{text}";
    }

    public static string formatSchemaName(string? prefix, string? name)
    {
      if (name == null)
        throw new Exception("'name' cannot be null");

      return addPrefix(prefix, name).Replace(".", "").ToLower();
    }

    public static string formatLogicalName(string? prefix, string? name)
    {
      if (name == null)
        throw new Exception("'name' cannot be null");

      return addPrefix(prefix, name).Replace(".", "").ToLower();
    }

    public static string formatLogicalCollectionName(string? prefix, string? name)
    {
      if (name == null)
        throw new Exception("'name' cannot be null");

      return $"{addPrefix(prefix, name).Replace(".", "").ToLower()}s".ToLower();
    }

    public static string formatDisplayName(string? prefix, string? name)
    {
      if (name == null)
        throw new Exception("'name' cannot be null");

      return name;
    }

    public static string formatDisplayCollectionName(string? prefix, string? name)
    {
      if (name == null)
        throw new Exception("'name' cannot be null");

      return $"{name}s";
    }

    public CreateEntityRequest generateEntityRequest(int locale, string? prefix, string? description, string? solutionUniqueName)
    {
      StringAttributeMetadata primaryAttribute = generatePrimaryAttribute(locale, prefix, description);

      string schemaName = formatSchemaName(prefix, Name);
      string displayName = formatDisplayName(prefix, Name);
      string displayCollectionName = formatDisplayCollectionName(prefix, Name);
      string logicalName = formatLogicalName(prefix, Name);
      string logicalCollectionName = formatLogicalCollectionName(prefix, Name);

      EntityMetadata entity = new()
      {
        SchemaName = schemaName,
        LogicalName = logicalName,
        LogicalCollectionName = logicalCollectionName,
        DisplayName = new Label(displayName, locale),
        DisplayCollectionName = new Label(displayCollectionName, locale),
        Description = new Label(description, locale),
        OwnershipType = OwnershipTypes.UserOwned,
        IsActivity = false,
      };

      // Create entity
      return new CreateEntityRequest()
      {
        Entity = entity,
        HasActivities = false,
        HasFeedback = false,
        HasNotes = false,
        PrimaryAttribute = primaryAttribute,
        SolutionUniqueName = solutionUniqueName
      };
    }

    private AttributeMetadata generateAttribute(int locale, string? prefix, string? description, KeyValuePair<string, string> member)
    {
      string schemaName = formatSchemaName(prefix, member.Key);
      string displayName = formatDisplayName(prefix, member.Key);
      string logicalName = formatLogicalName(prefix, member.Key);

      switch (member.Value.ToLower())
      {
        case "enum":
        case "string":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Text,
            MaxLength = 200
          };

        case "text":
          return new MemoAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.TextArea,
            MaxLength = 50000
          };

        case "richtext":
          return new MemoAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.RichText,
            MaxLength = 50000
          };

        case "email":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Email,
            MaxLength = 200
          };

        case "phone":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Phone,
            MaxLength = 200
          };

        case "url":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Url,
            MaxLength = 200
          };

        case "json":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Json,
            MaxLength = 200
          };

        case "boolean":
          return new BooleanAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            OptionSet = new BooleanOptionSetMetadata(
              new OptionMetadata(new Label("True", locale), 1),
              new OptionMetadata(new Label("False", locale), 0)
            )
          };

        case "datetime":
          return new DateTimeAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = DateTimeFormat.DateAndTime,
            ImeMode = ImeMode.Auto
          };

        case "integer":
          return new IntegerAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = IntegerFormat.None
          };

        case "long":
          return new BigIntAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None)
          };

        case "decimal":
          return new DecimalAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Precision = 10
          };

        case "binary":
          return new FileAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = new Label(displayName, locale),
            Description = new Label(description, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None)
          };

        default:
          throw new Exception($"Entity: {Name}, Member: {member.Key} - Data type value '{member.Value}' does not exist.");

      }
    }

    public List<CreateAttributeRequest> generateAttributeRequests(int locale, string? prefix, string? description, string? solutionUniqueName)
    {
      List<CreateAttributeRequest> requests = new();
      string entitySchemaName = formatSchemaName(prefix, Name);

      foreach (KeyValuePair<string, string> member in Members.Skip(1))
      {
        AttributeMetadata attribute = generateAttribute(locale, prefix, description, member);

        requests.Add(new()
        {
          Attribute = attribute,
          EntityName = entitySchemaName,
          SolutionUniqueName = solutionUniqueName
        });
      }

      return requests;
    }

    private StringAttributeMetadata generatePrimaryAttribute(int locale, string? prefix, string? description)
    {
      KeyValuePair<string, string> primaryMember = Members.First();

      if (primaryMember.Value.ToLower() != "string")
        throw new Exception($"Entity: {Name}, Member: {primaryMember.Key} - Invalid type '{primaryMember.Value}'. Primary attribute has to be a string.");

      return (StringAttributeMetadata)generateAttribute(locale, prefix, description, primaryMember);
    }

    public CreateEntityKeyRequest generateKeyRequest(int locale, string? prefix, string? solutionUniqueName)
    {
      string[] keyAttributes = { formatLogicalName(prefix, PrimaryKey) };

      string entityLogicalName = formatLogicalName(prefix, Name);

      string keySchemaName = formatSchemaName(prefix, PrimaryKey);
      string keyLogicalName = formatLogicalName(prefix, PrimaryKey);
      string keyDisplayName = formatDisplayName(prefix, PrimaryKey);


      return new CreateEntityKeyRequest()
      {
        EntityKey = new EntityKeyMetadata()
        {
          DisplayName = new Label(keyDisplayName, locale),
          LogicalName = keyLogicalName,
          SchemaName = keySchemaName,
          KeyAttributes = keyAttributes,
          IsSecondaryKey = false
        },
        EntityName = entityLogicalName,
        SolutionUniqueName = solutionUniqueName,
      };
    }

    public List<CreateOneToManyRequest> generateManyToOneRequests(int locale, string? prefix, string? solutionUniqueName)
    {
      List<CreateOneToManyRequest> requests = new();

      foreach (KeyValuePair<string, string> relationship in ManyToOne)
      {
        string referencedTableDisplayName = formatDisplayName(prefix, relationship.Value);
        string referencedTableLogicalName = formatLogicalName(prefix, relationship.Value);
        string referencingTableLogicalName = formatLogicalName(prefix, Name);

        string relationshipSchemaName = formatSchemaName(prefix, relationship.Key);
        string relationshipDisplayName = formatDisplayName(prefix, relationship.Key);
        string relationshipLogicalName = formatLogicalName(prefix, relationship.Key);

        requests.Add(new()
        {
          // Defines the lookup to create on the Referencing table
          Lookup = new LookupAttributeMetadata
          {
            SchemaName = relationshipSchemaName,
            LogicalName = relationshipLogicalName,
            DisplayName = new Label(relationshipDisplayName, locale),
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Description = new Label($"Lookup to the {referencedTableDisplayName} table.", locale)
          },
          // Defines the relationship to create to support the lookup
          OneToManyRelationship = new OneToManyRelationshipMetadata
          {
            ReferencedEntity = referencedTableLogicalName,
            ReferencingEntity = referencingTableLogicalName,
            SchemaName = relationshipSchemaName,
            // Controls how the relationship appears in model-driven apps
            AssociatedMenuConfiguration = new AssociatedMenuConfiguration
            {
              Behavior = AssociatedMenuBehavior.UseLabel,
              Group = AssociatedMenuGroup.Details,
              Label = new Label(relationshipDisplayName, locale),
              Order = 10000
            },
            // Controls automated behaviors for related records
            CascadeConfiguration = new CascadeConfiguration
            {
              Assign = CascadeType.NoCascade,
              Delete = CascadeType.RemoveLink,
              Merge = CascadeType.NoCascade,
              Reparent = CascadeType.NoCascade,
              Share = CascadeType.NoCascade,
              Unshare = CascadeType.NoCascade
            }
          },
          SolutionUniqueName = solutionUniqueName
        });
      }

      return requests;
    }

    public List<CreateManyToManyRequest> generateManyToManyRequests(int locale, string? prefix, string? solutionUniqueName)
    {
      List<CreateManyToManyRequest> requests = new();

      foreach (KeyValuePair<string, string> relationship in ManyToMany)
      {
        string referencedTableLogicalName = formatLogicalName(prefix, Name);
        string referencingTableLogicalName = formatLogicalName(prefix, relationship.Value);

        string relationshipSchemaName = formatSchemaName(prefix, relationship.Key);
        string relationshipDisplayName = formatDisplayName(prefix, relationship.Key);

        requests.Add(new()
        {
          IntersectEntitySchemaName = relationshipSchemaName,
          ManyToManyRelationship = new ManyToManyRelationshipMetadata()
          {
            SchemaName = relationshipSchemaName,
            Entity1LogicalName = referencedTableLogicalName,
            Entity1AssociatedMenuConfiguration = new AssociatedMenuConfiguration()
            {
              Behavior = AssociatedMenuBehavior.UseLabel,
              Group = AssociatedMenuGroup.Details,
              Label = new Label(relationshipDisplayName, locale),
              Order = 10000

            },
            Entity2LogicalName = referencingTableLogicalName,
            Entity2AssociatedMenuConfiguration = new AssociatedMenuConfiguration()
            {
              Behavior = AssociatedMenuBehavior.UseLabel,
              Group = AssociatedMenuGroup.Details,
              Label = new Label(relationshipDisplayName, locale),
              Order = 10000
            }
          },
          SolutionUniqueName = solutionUniqueName
        });
      }

      return requests;
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
    // Default parameters
    int defaultLocale = 1033; // 1033 or 2057 - for English
    string defaultPrefix = "OAP";
    string defaultDescription = "Imported from the Ops App.";
    string defaultSolution = "OAP";


    // Load JSON file here
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write("Enter file path to *.json file: ");
    Console.ResetColor();
    string? inputFilepath = Console.ReadLine();
    if (inputFilepath == null)
      throw new Exception("File path is required");

    inputFilepath = inputFilepath.Replace("\"", "");
    StreamReader stream = new StreamReader(inputFilepath);
    List<Excel2Dataverse.JSONEntity>? jsonEntities = JsonSerializer.Deserialize<List<Excel2Dataverse.JSONEntity>>(stream.ReadToEnd());

    if (jsonEntities == null)
      throw new Exception("No JSON to load - Invalid JSON.");


    // Get user input
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write("Enter a prefix (default 'OAP'): ");
    Console.ResetColor();
    string? inputPrefix = Console.ReadLine();

    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write("Enter a description (default 'Imported from the Ops App.'): ");
    Console.ResetColor();
    string? inputDescription = Console.ReadLine();

    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write("Enter solution unique name (default 'OAP'): ");
    Console.ResetColor();
    string? inputSolution = Console.ReadLine();

    defaultPrefix = inputPrefix == null || inputPrefix.Trim() == "" ? defaultPrefix : inputPrefix;
    defaultDescription = inputDescription == null || inputDescription.Trim() == "" ? defaultDescription : inputDescription;
    defaultSolution = inputSolution == null || inputSolution.Trim() == "" ? defaultSolution : inputSolution;

    Console.WriteLine("\n");


    // Login to Microsoft
    Console.WriteLine("Sign-in to your Microsoft Account (browser pop-up window) ...");
    IOrganizationService service = new ServiceClient(connectionString());
    var whoAmI = (WhoAmIResponse)service.Execute(new WhoAmIRequest());
    Console.WriteLine($"Successfully signed in. User ID is {whoAmI.UserId}.\n\n");

    // Load existing entities - to check if entities/attributes/relationships already exists
    Console.WriteLine("Retrieving all current entities from Dataverse ...\n\n");
    RetrieveAllEntitiesResponse retrieveAllEntities = (RetrieveAllEntitiesResponse)service.Execute(new RetrieveAllEntitiesRequest()
    {
      EntityFilters = EntityFilters.All,
      RetrieveAsIfPublished = false,
    });
    List<EntityMetadata> existingEntities = retrieveAllEntities.EntityMetadata.ToList();


    // Create entities
    Console.WriteLine($"Adding {jsonEntities.Count} {(jsonEntities.Count == 1 ? "entity" : "entities")} from '{inputFilepath}' ...");

    int entitiyCount = 1;
    foreach (Excel2Dataverse.JSONEntity jsonEntity in jsonEntities)
    {
      CreateEntityRequest entityRequest = jsonEntity.generateEntityRequest(defaultLocale, defaultPrefix, defaultDescription, defaultSolution);
      EntityMetadata? existingEntity = existingEntities.Find(x => x.LogicalName == entityRequest.Entity.LogicalName);

      List<CreateAttributeRequest> attributeRequests = jsonEntity.generateAttributeRequests(defaultLocale, defaultPrefix, defaultDescription, defaultSolution);
      List<AttributeMetadata> existingAttributes = existingEntity != null ? existingEntity.Attributes.ToList() : new();

      Console.Write($"\t[{entitiyCount}/{jsonEntities.Count}] - ");

      if (existingEntity != null)
      {
        Console.WriteLine($"Skipping '{entityRequest.Entity.SchemaName}' entity, already exists.");
      }
      else
      {
        service.Execute(entityRequest);
        Console.WriteLine($"'{entityRequest.Entity.SchemaName}' entity has been created.");
      }

      Console.WriteLine($"\tAdding {attributeRequests.Count} attribute{(attributeRequests.Count == 1 ? "" : "s")} ...");

      // Add attributes
      int attributeCount = 1;
      foreach (CreateAttributeRequest attributeRequest in attributeRequests)
      {
        AttributeMetadata? existingAttribute = existingAttributes.Find(x => x.LogicalName == attributeRequest.Attribute.LogicalName);

        Console.Write($"\t\t[{attributeCount}/{attributeRequests.Count}] - ");

        if (existingAttribute != null)
        {
          Console.WriteLine($"Skipping '{attributeRequest.Attribute.SchemaName}' attribute, already exists.");
        }
        else
        {
          service.Execute(attributeRequest);
          Console.WriteLine($"'{attributeRequest.Attribute.SchemaName}' attribute has been added.");
        }
        attributeCount++;
      }

      if (!(jsonEntity.PrimaryKey == null || jsonEntity.PrimaryKey.Trim() == ""))
      {
        Console.WriteLine("\tAdding key ...");

        // Add key
        List<EntityKeyMetadata> existingKeys = existingEntity != null ? existingEntity.Keys.ToList() : new();

        CreateEntityKeyRequest keyRequest = jsonEntity.generateKeyRequest(defaultLocale, defaultPrefix, defaultSolution);
        EntityKeyMetadata? existingKey = existingKeys.Find(x => x.LogicalName == keyRequest.EntityKey.LogicalName);

        if (existingKey != null)
        {
          Console.WriteLine($"\t\tSkipping '{keyRequest.EntityKey.LogicalName}' key, already exists.");
        }
        else
        {
          service.Execute(keyRequest);
          Console.WriteLine($"\t\t'{keyRequest.EntityKey.LogicalName}' key had been added.");
        }
      }

      entitiyCount++;
    }


    // Create relationship - Many to one
    foreach (Excel2Dataverse.JSONEntity jsonEntity in jsonEntities)
    {
      if (jsonEntity.ManyToOne.Count < 1)
        continue;

      EntityMetadata? existingEntity = existingEntities.Find(x => x.LogicalName == Excel2Dataverse.JSONEntity.formatLogicalName(defaultPrefix, jsonEntity.Name));
      List<OneToManyRelationshipMetadata> existingManyToOnes = existingEntity != null ? existingEntity.ManyToOneRelationships.ToList() : new();

      List<CreateOneToManyRequest> manyToOneRequests = jsonEntity.generateManyToOneRequests(defaultLocale, defaultPrefix, defaultSolution);
      Console.WriteLine($"\nAdding {jsonEntity.ManyToOne.Count} many-to-one relationship{(jsonEntity.ManyToOne.Count == 1 ? "" : "s")} to '{Excel2Dataverse.JSONEntity.formatSchemaName(defaultPrefix, jsonEntity.Name)}' ...");

      int relationshipCount = 1;
      foreach (CreateOneToManyRequest manyToOneRequest in manyToOneRequests)
      {
        OneToManyRelationshipMetadata? existingManyToOne = existingManyToOnes.Find(x => x.SchemaName == manyToOneRequest.OneToManyRelationship.SchemaName);

        Console.Write($"\t[{relationshipCount}/{manyToOneRequests.Count}] - ");

        if (existingManyToOne != null)
        {
          Console.WriteLine($"Skipping '{manyToOneRequest.OneToManyRelationship.SchemaName}' many-to-one relationship, already exists.");
        }
        else
        {
          service.Execute(manyToOneRequest);
          Console.WriteLine($"'{manyToOneRequest.OneToManyRelationship.SchemaName}' many-to-one relationship has been added ({manyToOneRequest.OneToManyRelationship.ReferencedEntity} *-> {manyToOneRequest.OneToManyRelationship.ReferencingEntity}).");
        }
        relationshipCount++;
      }
    }


    // Create relationship - Many to many
    foreach (Excel2Dataverse.JSONEntity jsonEntity in jsonEntities)
    {
      if (jsonEntity.ManyToMany.Count < 1)
        continue;

      EntityMetadata? existingEntity = existingEntities.Find(x => x.LogicalName == Excel2Dataverse.JSONEntity.formatLogicalName(defaultPrefix, jsonEntity.Name));
      List<ManyToManyRelationshipMetadata> existingManyToManys = existingEntity != null ? existingEntity.ManyToManyRelationships.ToList() : new();

      Console.WriteLine($"\nAdding {jsonEntity.ManyToMany.Count} many-to-many relationship{(jsonEntity.ManyToMany.Count == 1 ? "" : "s")} to '{Excel2Dataverse.JSONEntity.formatSchemaName(defaultPrefix, jsonEntity.Name)}':");

      List<CreateManyToManyRequest> manyToManyRequests = jsonEntity.generateManyToManyRequests(defaultLocale, defaultPrefix, defaultSolution);

      int relationshipCount = 1;
      foreach (CreateManyToManyRequest manyToManyRequest in manyToManyRequests)
      {
        ManyToManyRelationshipMetadata? existingManyToMany = existingManyToManys.Find(x => x.SchemaName == manyToManyRequest.ManyToManyRelationship.SchemaName);

        Console.Write($"\t[{relationshipCount}/{manyToManyRequests.Count}] - ");

        if (existingManyToMany != null)
        {
          Console.WriteLine($"Skipping '{manyToManyRequest.ManyToManyRelationship.SchemaName}' many-to-many relationship, already exists.");
        }
        else
        {
          service.Execute(manyToManyRequest);
          Console.WriteLine($"'{manyToManyRequest.ManyToManyRelationship.SchemaName}' many-to-many relationship has been added ({manyToManyRequest.ManyToManyRelationship.Entity1LogicalName} *-* {manyToManyRequest.ManyToManyRelationship.Entity2LogicalName}).");
        }
        relationshipCount++;
      }
    }

    // Pause the console so it does not close.
    Console.WriteLine("Press the <Enter> key to exit.");
    Console.ReadLine();
  }
}

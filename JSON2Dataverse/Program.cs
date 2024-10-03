using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Text.Json;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace JSON2Dataverse
{
  class Action
  {
    public IOrganizationService Service;
    private List<EntityMetadata> Entities = new();
    public Action() => Service = new ServiceClient(GetConnectionString());
    public Action(string connectionString) => Service = new ServiceClient(connectionString);
    private static string GetConnectionString()
    {
      ConnectionStringSettings settings = ConfigurationManager.ConnectionStrings["dynamics"];
      return settings.ConnectionString;
    }

    public List<EntityMetadata> GetEntities()
    {
      RetrieveAllEntitiesResponse allEntities = (RetrieveAllEntitiesResponse)Service.Execute(new RetrieveAllEntitiesRequest()
      {
        EntityFilters = EntityFilters.All,
        RetrieveAsIfPublished = false,
      });

      Entities = allEntities.EntityMetadata.ToList();
      return Entities;
    }
    public void CreateEntities(List<JSONEntity> jsonEntities)
    {
      Console.WriteLine($"Adding {jsonEntities.Count} {(jsonEntities.Count == 1 ? "entity" : "entities")} ...");

      int entitiyCount = 1;
      foreach (JSONEntity jsonEntity in jsonEntities)
      {
        CreateEntityRequest entityRequest = jsonEntity.GenerateEntityRequest();
        EntityMetadata? existingEntity = Entities.Find(x => x.LogicalName == entityRequest.Entity.LogicalName);
        List<CreateAttributeRequest> attributeRequests = jsonEntity.GenerateAttributeRequests();

        List<AttributeMetadata> existingAttributes = existingEntity != null ? existingEntity.Attributes.ToList() : new();

        Console.Write($"\t[{entitiyCount}/{jsonEntities.Count}] - ");

        if (existingEntity != null)
        {
          Console.WriteLine($"Skipping '{entityRequest.Entity.SchemaName}' entity, already exists.");
        }
        else
        {
          Service.Execute(entityRequest);

          Console.ForegroundColor = ConsoleColor.Green;
          Console.WriteLine($"'{entityRequest.Entity.SchemaName}' entity has been created.");
          Console.ResetColor();
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
            Service.Execute(attributeRequest);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"'{attributeRequest.Attribute.SchemaName}' attribute has been added.");
            Console.ResetColor();
          }
          attributeCount++;
        }

        if (!(jsonEntity.PrimaryKey == null || jsonEntity.PrimaryKey.Trim() == ""))
        {
          Console.WriteLine("\tAdding key ...");

          // Add key
          List<EntityKeyMetadata> existingKeys = existingEntity != null ? existingEntity.Keys.ToList() : new();

          CreateEntityKeyRequest keyRequest = JSONEntity.GenerateKeyRequest(jsonEntity.Name, jsonEntity.PrimaryKey, new string[] { jsonEntity.PrimaryKey });
          EntityKeyMetadata? existingKey = existingKeys.Find(x => x.LogicalName == keyRequest.EntityKey.LogicalName);

          if (existingKey != null)
          {
            Console.WriteLine($"\t\tSkipping '{keyRequest.EntityKey.LogicalName}' key, already exists.");
          }
          else
          {
            Service.Execute(keyRequest);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\t\t'{keyRequest.EntityKey.LogicalName}' key had been added.");
            Console.ResetColor();
          }
        }

        entitiyCount++;
      }
    }
    public void CreateManyToOne(List<JSONEntity> jsonEntities)
    {
      foreach (JSONEntity jsonEntity in jsonEntities)
      {
        if (jsonEntity.ManyToOne.Count < 1)
          continue;

        EntityMetadata? existingEntity = Entities.Find(x => x.LogicalName == jsonEntity.LogicalName);
        List<OneToManyRelationshipMetadata> existingManyToOnes = existingEntity != null ? existingEntity.ManyToOneRelationships.ToList() : new();
        List<CreateOneToManyRequest> manyToOneRequests = jsonEntity.GenerateManyToOneRequests();

        Console.WriteLine($"\nAdding {jsonEntity.ManyToOne.Count} many-to-one relationship{(jsonEntity.ManyToOne.Count == 1 ? "" : "s")} to '{jsonEntity.SchemaName}' ...");

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
            Service.Execute(manyToOneRequest);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"'{manyToOneRequest.OneToManyRelationship.SchemaName}' many-to-one relationship has been added ({manyToOneRequest.OneToManyRelationship.ReferencedEntity} *-> {manyToOneRequest.OneToManyRelationship.ReferencingEntity}).");
            Console.ResetColor();
          }
          relationshipCount++;
        }
      }
    }
    public void CreateManyToMany(List<JSONEntity> jsonEntities)
    {
      foreach (JSONEntity jsonEntity in jsonEntities)
      {
        if (jsonEntity.ManyToMany.Count < 1)
          continue;

        EntityMetadata? existingEntity = Entities.Find(x => x.LogicalName == jsonEntity.LogicalName);
        List<ManyToManyRelationshipMetadata> existingManyToManys = existingEntity != null ? existingEntity.ManyToManyRelationships.ToList() : new();
        List<CreateManyToManyRequest> manyToManyRequests = jsonEntity.GenerateManyToManyRequests();

        Console.WriteLine($"\nAdding {jsonEntity.ManyToMany.Count} many-to-many relationship{(jsonEntity.ManyToMany.Count == 1 ? "" : "s")} to '{jsonEntity.SchemaName}':");

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
            Service.Execute(manyToManyRequest);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"'{manyToManyRequest.ManyToManyRelationship.SchemaName}' many-to-many relationship has been added ({manyToManyRequest.ManyToManyRelationship.Entity1LogicalName} *-* {manyToManyRequest.ManyToManyRelationship.Entity2LogicalName}).");
            Console.ResetColor();
          }
          relationshipCount++;
        }
      }
    }
    public void CreateManyToManyIntersection(List<JSONEntity> jsonEntities)
    {
      foreach (JSONEntity jsonEntity in jsonEntities)
      {
        if (jsonEntity.ManyToMany.Count < 1)
          continue;

        Console.WriteLine($"\nAdding {jsonEntity.ManyToMany.Count} many-to-many relationship{(jsonEntity.ManyToMany.Count == 1 ? "" : "s")} to '{jsonEntity.SchemaName}':");

        int relationshipCount = 1;
        foreach (KeyValuePair<string, string> manyToMany in jsonEntity.ManyToMany)
        {
          EntityMetadata? existingEntity = Entities.Find(x => x.SchemaName == JSONEntity.FormatSchemaName(JSONEntity.Prefix, manyToMany.Key));

          Console.Write($"\t[{relationshipCount}/{jsonEntity.ManyToMany.Count}] - ");

          if (existingEntity != null)
          {
            Console.WriteLine($"Skipping '{existingEntity.SchemaName}' many-to-many relationship, already exists.");
          }
          else
          {
            string lookup1Name = JSONEntity.AddPrefix(manyToMany.Key, jsonEntity.Name);
            string lookup2Name = JSONEntity.AddPrefix(manyToMany.Key, manyToMany.Value);

            // Intersection table
            var intersectionTable = JSONEntity.GenerateEntityRequest(manyToMany.Key, new()
            {
              SchemaName = JSONEntity.FormatSchemaName(JSONEntity.Prefix, "Name"),
              LogicalName = JSONEntity.FormatSchemaName(JSONEntity.Prefix, "Name"),
              DisplayName = new(JSONEntity.FormatDisplayName("Name"), JSONEntity.Locale),
              Description = JSONEntity.DisplayDescription,
              RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
              Format = StringFormat.Text,
              MaxLength = 200
            });

            // Lookup 1
            var lookup1 = JSONEntity.GenerateOneToManyRequest(jsonEntity.Name, manyToMany.Key, lookup1Name);

            // Lookup 2
            var lookup2 = JSONEntity.GenerateOneToManyRequest(manyToMany.Value, manyToMany.Key, lookup2Name);

            // Key
            var keyRequest = JSONEntity.GenerateKeyRequest(manyToMany.Key, "RelationshipID", new string[] { lookup1Name, lookup2Name });

            Service.Execute(intersectionTable);
            Service.Execute(lookup1);
            Service.Execute(lookup2);
            Service.Execute(keyRequest);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"'{intersectionTable.Entity.SchemaName}' many-to-many relationship has been added ({lookup1.OneToManyRelationship.ReferencedEntity} *-* {lookup2.OneToManyRelationship.ReferencedEntity}).");
            Console.ResetColor();
          }
          relationshipCount++;
        }
      }
    }
  }
  class JSONEntity
  {
    [Required] public string? Name { get; set; }
    [Required] public string? PrimaryKey { get; set; }
    public static int Locale { get; set; } = 1033; // 1033 or 2057 - for English
    public static string Prefix { get; set; } = "OAP";
    public static string Description { get; set; } = "Imported from the Ops App.";
    public static string Solution { get; set; } = "OAP";
    public Dictionary<string, string> Members { get; set; } = new();
    public Dictionary<string, string> ManyToOne { get; set; } = new();
    public Dictionary<string, string> ManyToMany { get; set; } = new();


    public string SchemaName => FormatSchemaName(Prefix, Name);
    public string LogicalName => FormatLogicalName(Prefix, Name);
    public string LogicalCollectionName => FormatLogicalCollectionName(Prefix, Name);
    public Label DisplayName => new(FormatDisplayName(Name), Locale);
    public Label DisplayCollectionName => new(FormatDisplayCollectionName(Name), Locale);
    public static Label DisplayDescription => new(Description, Locale);


    public static string AddPrefix(string? prefix, string? text)
    {
      if (text == null)
        throw new Exception("Argument 'text' cannot be null");

      if (prefix == null || prefix.Trim() == "")
        return text;

      return $"{prefix}_{text}";
    }
    public static string FormatSchemaName(string? prefix, string? name)
    {
      if (name == null)
        throw new Exception("Argument 'name' cannot be null");

      return AddPrefix(prefix, name).Replace(".", "").ToLower();
    }
    public static string FormatLogicalName(string? prefix, string? name)
    {
      if (name == null)
        throw new Exception("Argument 'name' cannot be null");

      return AddPrefix(prefix, name).Replace(".", "").ToLower();
    }
    public static string FormatLogicalCollectionName(string? prefix, string? name)
    {
      if (name == null)
        throw new Exception("Argument 'name' cannot be null");

      return $"{AddPrefix(prefix, name).Replace(".", "").ToLower()}s".ToLower();
    }
    public static string FormatDisplayName(string? name)
    {
      if (name == null)
        throw new Exception("Argument 'name' cannot be null");

      if (name == null)
        throw new Exception("Argument 'name' cannot be null");

      return name;
    }
    public static string FormatDisplayCollectionName(string? name)
    {
      if (name == null)
        throw new Exception("Argument 'name' cannot be null");

      return $"{name}s";
    }


    public static AttributeMetadata GenerateAttribute(string? attributeName, string datatype)
    {
      if (attributeName == null)
        throw new Exception("Argument 'attributeName' cannot be null.");

      string schemaName = FormatSchemaName(Prefix, attributeName);
      string logicalName = FormatLogicalName(Prefix, attributeName);
      Label displayName = new(FormatDisplayName(attributeName), Locale);
      Label descriptionLabel = DisplayDescription;

      switch (datatype.ToLower())
      {
        case "enum":
        case "string":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Text,
            MaxLength = 200
          };

        case "text":
          return new MemoAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.TextArea,
            MaxLength = 50000
          };

        case "richtext":
          return new MemoAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.RichText,
            MaxLength = 50000
          };

        case "email":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Email,
            MaxLength = 200
          };

        case "phone":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Phone,
            MaxLength = 200
          };

        case "url":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Url,
            MaxLength = 200
          };

        case "json":
          return new StringAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = StringFormat.Json,
            MaxLength = 200
          };

        case "boolean":
          return new BooleanAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            OptionSet = new BooleanOptionSetMetadata(
              new OptionMetadata(new Label("True", Locale), 1),
              new OptionMetadata(new Label("False", Locale), 0)
            )
          };

        case "datetime":
          return new DateTimeAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = DateTimeFormat.DateAndTime,
            ImeMode = ImeMode.Auto
          };

        case "integer":
          return new IntegerAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Format = IntegerFormat.None
          };

        case "long":
          return new BigIntAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None)
          };

        case "decimal":
          return new DecimalAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
            Precision = 10
          };

        case "binary":
          return new FileAttributeMetadata()
          {
            SchemaName = schemaName,
            LogicalName = logicalName,
            DisplayName = displayName,
            Description = descriptionLabel,
            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None)
          };

        default:
          throw new Exception($" Member: {attributeName} - Data type value '{datatype}' does not exist.");

      }
    }
    public static CreateOneToManyRequest GenerateOneToManyRequest(string? referencedTable, string? referencingTable, string relationshipName)
    { 
      if (referencedTable == null)
        throw new Exception("Argument 'referencedTable' cannot be null.");
      if (referencingTable == null)
        throw new Exception("Argument 'referencingTable' cannot be null.");

      string referencedTableLogicalName = FormatLogicalName(Prefix, referencedTable);
      string referencingTableLogicalName = FormatLogicalName(Prefix, referencingTable);
      Label referencedTableDisplayName = new(FormatDisplayName(referencedTable), Locale);

      string relationshipSchemaName = FormatSchemaName(Prefix, relationshipName);
      string relationshipLogicalName = FormatLogicalName(Prefix, relationshipName);
      Label relationshipDisplayName = new(FormatDisplayName(relationshipName), Locale);

      return new()
      {
        // Defines the lookup to create on the Referencing table
        Lookup = new LookupAttributeMetadata
        {
          SchemaName = relationshipSchemaName,
          LogicalName = relationshipLogicalName,
          DisplayName = relationshipDisplayName,
          RequiredLevel = new(AttributeRequiredLevel.None),
          Description = new($"Lookup to the {referencedTableDisplayName} table.", Locale)
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
            Label = relationshipDisplayName,
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
        SolutionUniqueName = Solution
      };
    }
    public static CreateEntityRequest GenerateEntityRequest(string? entityName, StringAttributeMetadata primaryAttribute)
    {
      if (entityName == null)
        throw new Exception("Argument 'entityName' cannot be null.");

      EntityMetadata entity = new()
      {
        SchemaName = FormatSchemaName(Prefix, entityName),
        LogicalName = FormatLogicalName(Prefix, entityName),
        LogicalCollectionName = FormatLogicalCollectionName(Prefix, entityName),
        DisplayName = new(FormatDisplayName(entityName), Locale),
        DisplayCollectionName = new(FormatDisplayCollectionName(entityName), Locale),
        Description = DisplayDescription,
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
        SolutionUniqueName = Solution
      };
    }
    public static CreateEntityKeyRequest GenerateKeyRequest(string? entityName, string keyName, string[] keyAttributeNames)
    {
      if (entityName == null)
        throw new Exception("Argument 'entityName' cannot be null.");

      string[] keyAttributes = keyAttributeNames.Select(attributeName => FormatLogicalName(Prefix, attributeName)).ToArray();

      string entityLogicalName = FormatLogicalName(Prefix, entityName);

      string keySchemaName = FormatSchemaName(Prefix, keyName);
      string keyLogicalName = FormatLogicalName(Prefix, keyName);
      Label keyDisplayName = new(FormatDisplayName(keyName), Locale);


      return new CreateEntityKeyRequest()
      {
        EntityKey = new EntityKeyMetadata()
        {
          DisplayName = keyDisplayName,
          LogicalName = keyLogicalName,
          SchemaName = keySchemaName,
          KeyAttributes = keyAttributes,
          IsSecondaryKey = false
        },
        EntityName = entityLogicalName,
        SolutionUniqueName = Solution,
      };
    }


    private StringAttributeMetadata GeneratePrimaryAttribute()
    {
      KeyValuePair<string, string> primaryMember = Members.First();

      if (primaryMember.Value.ToLower() != "string")
        throw new Exception($"Entity: {Name}, Member: {primaryMember.Key} - Invalid type '{primaryMember.Value}'. Primary attribute has to be a string.");

      return (StringAttributeMetadata)GenerateAttribute(primaryMember.Key, primaryMember.Value);
    }


    public CreateEntityRequest GenerateEntityRequest() => GenerateEntityRequest(Name, GeneratePrimaryAttribute());
    public List<CreateAttributeRequest> GenerateAttributeRequests()
    {
      List<CreateAttributeRequest> requests = new();

      foreach (KeyValuePair<string, string> member in Members.Skip(1))
      {
        AttributeMetadata attribute = GenerateAttribute(member.Key, member.Value);

        requests.Add(new()
        {
          Attribute = attribute,
          EntityName = SchemaName,
          SolutionUniqueName = Solution
        });
      }

      return requests;
    }
    public List<CreateOneToManyRequest> GenerateManyToOneRequests()
    {
      if (Name == null)
        throw new Exception("Property 'Name' cannot be null.");

      List<CreateOneToManyRequest> requests = new();

      foreach (KeyValuePair<string, string> relationship in ManyToOne)
        requests.Add(GenerateOneToManyRequest(relationship.Value, Name, relationship.Key));

      return requests;
    }
    public List<CreateManyToManyRequest> GenerateManyToManyRequests()
    {
      List<CreateManyToManyRequest> requests = new();

      foreach (KeyValuePair<string, string> relationship in ManyToMany)
      {
        string referencedTableLogicalName = FormatLogicalName(Prefix, Name);
        string referencingTableLogicalName = FormatLogicalName(Prefix, relationship.Value);

        string relationshipSchemaName = FormatSchemaName(Prefix, relationship.Key);
        string relationshipDisplayName = FormatDisplayName(relationship.Key);

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
              Label = new(relationshipDisplayName, Locale),
              Order = 10000

            },
            Entity2LogicalName = referencingTableLogicalName,
            Entity2AssociatedMenuConfiguration = new AssociatedMenuConfiguration()
            {
              Behavior = AssociatedMenuBehavior.UseLabel,
              Group = AssociatedMenuGroup.Details,
              Label = new(relationshipDisplayName, Locale),
              Order = 10000
            }
          },
          SolutionUniqueName = Solution
        });
      }

      return requests;
    }

  }
}

class Program
{
  static void Main()
  {
    // Load JSON file here
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write("Enter file path to *.json file: ");
    Console.ResetColor();

    string? inputFilepath = Console.ReadLine();
    if (inputFilepath == null)
      throw new Exception("File path is required");

    inputFilepath = inputFilepath.Replace("\"", "");

    StreamReader stream = new StreamReader(inputFilepath);
    List<JSON2Dataverse.JSONEntity>? jsonEntities = JsonSerializer.Deserialize<List<JSON2Dataverse.JSONEntity>>(stream.ReadToEnd());

    if (jsonEntities == null)
      throw new Exception("No JSON to load - Invalid JSON.");


    // Get user input
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write($"Enter a prefix (default '{JSON2Dataverse.JSONEntity.Prefix}'): ");
    Console.ResetColor();
    string? inputPrefix = Console.ReadLine();

    if (inputPrefix != null && inputPrefix.Trim() != "")
      JSON2Dataverse.JSONEntity.Prefix = inputPrefix;

    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write($"Enter a description (default '{JSON2Dataverse.JSONEntity.Description}'): ");
    Console.ResetColor();
    string? inputDescription = Console.ReadLine();

    if (inputDescription != null && inputDescription.Trim() != "")
      JSON2Dataverse.JSONEntity.Description = inputDescription;

    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write($"Enter solution unique name (default '{JSON2Dataverse.JSONEntity.Solution}'): ");
    Console.ResetColor();
    string? inputSolution = Console.ReadLine();

    if (inputSolution != null && inputSolution.Trim() != "")
      JSON2Dataverse.JSONEntity.Solution = inputSolution;

    Console.WriteLine("\n");


    // Login to Microsoft
    Console.WriteLine("Sign-in to your Microsoft Account using the browser pop-up ...");
    JSON2Dataverse.Action action = new();
    Console.WriteLine($"Successfully signed in.\n\n");


    // Load existing entities - to check if entities/attributes/relationships already exists
    Console.WriteLine("Retrieving all current entities from Dataverse ...\n\n");
    action.GetEntities();


    // Create entities
    action.CreateEntities(jsonEntities);


    // Create relationship - Many to one
    action.CreateManyToOne(jsonEntities);


    // Create relationship - Many to many
    action.CreateManyToManyIntersection(jsonEntities);

    // Pause the console so it does not close.
    Console.WriteLine("Press the <Enter> key to exit.");
    Console.ReadLine();
  }
}

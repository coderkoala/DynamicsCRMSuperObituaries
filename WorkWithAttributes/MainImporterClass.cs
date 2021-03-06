﻿using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Text;
using System.Collections.Generic;
using System.Threading;

namespace PowerApps.ImporterClassMaster
{
    public partial class AttributesFactoryClass
    {
        static void clean()
        {
            Console.Clear();
        }

  /// <summary>
        ///
        /// </summary>
        /// <param name="SchemaName">Schema Name</param>
        /// <param name="DisplayName">Display Name</param>
        /// <param name="StringFormat">StringFormat.Url, StringFormat.Phone, StringFormat.Email</param>
        /// <param name="addedAttributes">Pass by reference.</param>
        /// <param name="lines">Pick among theses: 'singleLine', 'money', 'multi'. The option 'multi' is for memo.</param>
        static void createFieldString(string SchemaName, string DisplayName, StringFormat StringFormat, ref List<AttributeMetadata> addedAttributes, string lines = "singleLine")
        {

            if ("singleLine" == lines)
            {
                var EntityAttribute = new StringAttributeMetadata("new_" + SchemaName)
                {
                    SchemaName = "new_" + SchemaName,
                    LogicalName = "new_" + SchemaName,
                    DisplayName = new Label(DisplayName + " *", 1033),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Description = new Label("MSVProperties CRM - " + DisplayName, 1033),
                    MaxLength = 255,
                    IsValidForForm = true,
                    IsValidForGrid = true,
                    Format = StringFormat,
                };
                addedAttributes.Add(EntityAttribute);
            }
            else if ("money" == lines)
            {
                var EntityAttribute = new MoneyAttributeMetadata("new_" + SchemaName)
                {
                    SchemaName = "new_" + SchemaName,
                    LogicalName = "new_" + SchemaName,
                    DisplayName = new Label(DisplayName + " *", 1033),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Description = new Label("MSVProperties CRM - " + DisplayName, 1033),
                    // Set extended properties
                    MinValue = 0.00,
                    Precision = 2,
                    PrecisionSource = 1,
                    ImeMode = ImeMode.Disabled,
                    IsValidForForm = true,
                    IsValidForGrid = true,
                };

                addedAttributes.Add(EntityAttribute);
            }
            else
            {
                var EntityAttribute = new MemoAttributeMetadata("new_" + SchemaName)
                {
                    SchemaName = "new_" + SchemaName,
                    LogicalName = "new_" + SchemaName,
                    DisplayName = new Label(DisplayName + " *", 1033),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Description = new Label("MSVProperties CRM - " + DisplayName, 1033),
                    // Set extended properties
                    Format = StringFormat.TextArea,
                    ImeMode = ImeMode.Disabled,
                    MaxLength = 1000,
                    IsValidForForm = true,
                    IsValidForGrid = true,
                };
                addedAttributes.Add(EntityAttribute);
            }
        }

        /// <summary>
        /// Create Picklist Field
        /// </summary>
        /// <param name="SchemaName">Schema name of Attribute, as well as the optionset</param>
        /// <param name="DisplayName">Display name of Attribute.</param>
        /// <param name="OptionSetName">Option set to be mapped to(string).</param>
        /// <param name="addedAttributes">Pass by reference, your Entity List.</param>
        /// <param name="multi">Pass "multi" for multiple picklist</param>
        static void createFieldPicklist(string SchemaName, string DisplayName, string OptionSetName, ref List<AttributeMetadata> addedAttributes, string multi = "single")
        {
            if ("multi" == multi)
            {
                var CreatedMultiSelectPicklistAttributeMetadata = new MultiSelectPicklistAttributeMetadata()
                {
                    SchemaName = "new_" + SchemaName,
                    LogicalName = "new_" + SchemaName,
                    DisplayName = new Label(DisplayName + " *", 1033),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Description = new Label("MSVProperties CRM " + DisplayName + " multi check List", 1033),
                    IsValidForForm = true,
                    IsValidForGrid = true,
                    OptionSet = new OptionSetMetadata
                    {
                        IsGlobal = false,
                        Name = OptionSetName
                    }
                };

                // Add and return early.
                addedAttributes.Add(CreatedMultiSelectPicklistAttributeMetadata);
            }
            else
            {
                var CreatedPicklistAttributeMetadata = new PicklistAttributeMetadata()
                {
                    SchemaName = "new_" + SchemaName,
                    LogicalName = "new_" + SchemaName,
                    DisplayName = new Label(DisplayName + " *", 1033),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Description = new Label("MSVProperties CRM  " + DisplayName + " single checklist", 1033),
                    IsValidForForm = true,
                    IsValidForGrid = true,
                    OptionSet = new OptionSetMetadata
                    {
                        IsGlobal = false,
                        Name = OptionSetName
                    }
                };
                addedAttributes.Add(CreatedPicklistAttributeMetadata);
            }

        }

        /// <summary>
        /// Creates Boolean field for the Entity.
        /// </summary>
        /// <param name="SchemaName">The Schema Name</param>
        /// <param name="DisplayName">The field Display Name</param>
        /// <param name="trueValue">Truth Value Label</param>
        /// <param name="falseValue">False Value Label</param>
        /// <param name="addedAttributes">Pass by reference (ref addedAttributes)</param>
        /// <returns></returns>
        static void createFieldBoolean(string SchemaName, string DisplayName, string trueValue, string falseValue, ref List<AttributeMetadata> addedAttributes)
        {
            var CreatedBooleanAttributeMetadata = new BooleanAttributeMetadata()
            {
                SchemaName = "new_" + SchemaName,
                LogicalName = "new_" + SchemaName,
                DisplayName = new Label(DisplayName + "*", _languageCode),
                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                Description = new Label("MSVProperties CRM - " + DisplayName, _languageCode),
                IsValidForForm = true,
                IsValidForGrid = true,
                OptionSet = new BooleanOptionSetMetadata(
                    new OptionMetadata(new Label(trueValue, 1033), 1),
                    new OptionMetadata(new Label(falseValue, 1033), 0)
                ),
            };
            addedAttributes.Add(CreatedBooleanAttributeMetadata);
        }


        /// <summary>
        /// Adds the Date(or Datetime) field to the given entity.
        /// </summary>
        /// </summary>
        /// <param name="SchemaName">Entity Schema Name</param>
        /// <param name="DisplayName">Entity Display name</param>
        /// <param name="pickListArray">Array of strings for optionset(string)</param>
        /// <param name="addedAttributes">Pass by reference(the entity List object)</param>
        /// <param name="format">DateFormat.DateOnly, DateFormat.DateOnly</param>
        /// <returns></returns>
        static void createFieldDate(string SchemaName, string DisplayName, ref List<AttributeMetadata> addedAttributes, DateTimeFormat format)
        {
            var CreatedDateTimeAttributeMetadata = new DateTimeAttributeMetadata
            {
                SchemaName = "new_" + SchemaName,
                LogicalName = "new_" + SchemaName,
                DisplayName = new Label(DisplayName + "*", _languageCode),
                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                Description = new Label("MSVProperties CRM - " + DisplayName + "Multi Checklist", _languageCode),
                // Set extended properties
                Format = format,
                ImeMode = ImeMode.Disabled,
                IsValidForForm = true,
                IsValidForGrid = true,
            };

            addedAttributes.Add(CreatedDateTimeAttributeMetadata);
        }

        //[STAThread]
        /// <summary>
        /// Main function for execution Service Request
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            CrmServiceClient service = null;
            try
            {
                addedAttributes = new List<AttributeMetadata>();
                service = SampleHelpers.Connect("Connect");
                if (service.IsReady)
                {
                    #region FieldImport
                    #endregion FieldImport
                    List<string> attributesnotAdded = new List<string>();
                    int attributesCount = addedAttributes.Count;
                    int i = 1; // Counter.
                    foreach (AttributeMetadata addedAttributesTuples in addedAttributes)
                    {
                        // Create the request.
                        CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                        {
                            EntityName = "@todo",
                            Attribute = addedAttributesTuples
                        };

                        // Execute the request.
                        try
                        {
                            Console.WriteLine("[{0}/{1}] Attempting upserting attribute : " + addedAttributesTuples.SchemaName, i, attributesCount);
                            service.Execute(createAttributeRequest);
                            Console.WriteLine("[{0}/{1}] Success upserting attribute : " + addedAttributesTuples.SchemaName, i, attributesCount);
                        }
                        catch (Exception ex)
                        {
                            // Supress error.
                            attributesnotAdded.Add(addedAttributesTuples.SchemaName);
                            Console.WriteLine("[{0}/{1}] Failed upserting attribute : " + addedAttributesTuples.SchemaName, i, attributesCount);
                        }
                        ++i;
                    }
                    // NOTE: All customizations must be published before they can be used.
                    service.Execute(new PublishAllXmlRequest());

                    // Reporting begins here.
                    Console.WriteLine("Published all customizations for MSVproperties for the <{0}> entity.", "@todo");

                    // For Attributes that failed.
                    Console.WriteLine("Attributes that failed migrating\n===\n");
                    foreach (dynamic fieldname in attributesnotAdded)
                    {
                        Console.WriteLine("* new_" + fieldname);
                    }
                    Console.ReadLine();
                    CleanUpSample(service);
                }
                else
                {
                    const string UNABLE_TO_LOGIN_ERROR = "Unable to Login to Common Data Service";
                    if (service.LastCrmError.Equals(UNABLE_TO_LOGIN_ERROR))
                    {
                        Console.WriteLine("Check the connection string values in cds/App.config.");
                        throw new Exception(service.LastCrmError);
                    }
                    else
                    {
                        throw service.LastCrmException;
                    }
                }
            }
            catch (Exception ex)
            {
                //Supress Error. Just call the stack if you wanna find out what exactly went wrong.
                Console.WriteLine("Press Enter to continue");
                Console.ReadLine();
            }

            finally
            {
                if (service != null)
                    service.Dispose();
            }

        }
    }
}

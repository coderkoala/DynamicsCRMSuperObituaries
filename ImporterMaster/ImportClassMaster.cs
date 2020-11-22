using Microsoft.Crm.Sdk.Messages;
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
                    MinValue = -99999999999.00,
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
        /// <param name="addedAttributes">Pass by reference, your Entity List.</param>
        /// <param name="multi">Pass "multi" for multiple picklist</param>
        static void createFieldPicklist(string SchemaName, string DisplayName, ref List<AttributeMetadata> addedAttributes, string multi = "single")
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
                        OptionSetType = OptionSetType.Picklist,
                        Name = SchemaName + "_local",
                        Options = {
                                new OptionMetadata( new Label("-", _languageCode), null),
                        }
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
                        Name = SchemaName + "_local",
                        Options = {
                                new OptionMetadata( new Label("-", _languageCode), null),
                        }
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
                    createFieldString("SuperObitID", "Super Obit ID", StringFormat.Text, ref addedAttributes, "singleLine");
                    createFieldString("SuperObitName", "Super Obit Name", StringFormat.Text, ref addedAttributes, "singleLine");
                    createFieldString("Email", "Email", StringFormat.Email, ref addedAttributes, "singleLine");
                    createFieldString("Secondary Email", "Secondary Email", StringFormat.Email, ref addedAttributes, "singleLine");
                    createFieldString("ZestimateValue", "Zestimate Value", StringFormat.Text, ref addedAttributes, "money");
                    createFieldDate("DateofDeath", "Date of Death", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldString("RPRValue", "RPR Value", StringFormat.Text, ref addedAttributes, "money");
                    createFieldDate("DatePosted", "Date Posted", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldBoolean("HasMortgage", "Has Mortgage?", "Yes", "No", ref addedAttributes);
                    createFieldString("FuneralName", "Funeral Name", StringFormat.Text, ref addedAttributes);
                    createFieldString("Realtor_com", "Realtor.com Estimate", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("FuneralAddress", "Funeral Address", StringFormat.Text, ref addedAttributes);
                    createFieldString("RealtytracValue", "Realtytrac Value", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("FuneralContactNumber", "Funeral Contact Number", StringFormat.Phone, ref addedAttributes, "singleLine");
                    createFieldString("Summary", "Summary", StringFormat.Text, ref addedAttributes, "multi");
                    createFieldString("Age", "Age", StringFormat.Text, ref addedAttributes, "singleLine");
                    createFieldString("TotalLoans", "Total Loans", StringFormat.Text, ref addedAttributes, "money");
                    createFieldBoolean("Deceased", "Deceased?", "Yes", "No", ref addedAttributes);
                    createFieldString("DeceasedLastAddress", "Deceased Last Address", StringFormat.Text, ref addedAttributes);
                    createFieldString("ObitLink", "Obit Link", StringFormat.Url, ref addedAttributes);
                    createFieldString("DeceasedLastCity", "Deceased Last City", StringFormat.Text, ref addedAttributes);
                    createFieldString("PlaceofDeath", "Place of Death", StringFormat.Text, ref addedAttributes);
                    createFieldString("DeceasedLastState", "Deceased Last State", StringFormat.Text, ref addedAttributes);
                    createFieldDate("DateofBirth", "Date of Birth", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldPicklist("Category", "Category", ref addedAttributes);
                    createFieldString("DeceasedLastZip", "Deceased Last Zip", StringFormat.Text, ref addedAttributes);
                    createFieldString("DeceasedFirstName", "Deceased First Name", StringFormat.Text, ref addedAttributes);
                    createFieldString("County", "County", StringFormat.Text, ref addedAttributes);
                    createFieldString("DeceasedMiddleName", "Deceased Middle Name", StringFormat.Text, ref addedAttributes);
                    createFieldString("NumberofSon", "Number of Son", StringFormat.Text, ref addedAttributes);
                    createFieldString("DeceasedLastName", "Deceased Last Name", StringFormat.Text, ref addedAttributes);
                    createFieldString("NumberofDaughter", "Number of Daughter", StringFormat.Text, ref addedAttributes);
                    createFieldPicklist("SalesStage", "Sales Stage", ref addedAttributes);
                    createFieldPicklist("Status", "Status", ref addedAttributes);
                    createFieldString("MainContactNumber", "Main Contact Number", StringFormat.Phone, ref addedAttributes);
                    createFieldPicklist("PropertyStatus", "Property Status", ref addedAttributes);
                    createFieldString("MainContactPerson", "Main Contact Person", StringFormat.Text, ref addedAttributes);
                    createFieldPicklist("Relationship", "Relationship", ref addedAttributes);
                    createFieldString("MainContactAddress", "Main Contact Address", StringFormat.Text, ref addedAttributes);
                    createFieldString("MainContactCity", "Main Contact City", StringFormat.Text, ref addedAttributes);
                    createFieldString("MainContactState", "Main Contact State", StringFormat.Text, ref addedAttributes);
                    createFieldString("MainContactZipcode", "Main Contact Zipcode", StringFormat.Text, ref addedAttributes);
                    createFieldString("LeadNumber", "Lead Number", StringFormat.Text, ref addedAttributes);
                    createFieldString("DeceasedFullName", "Deceased Full Name", StringFormat.Text, ref addedAttributes);
                    createFieldPicklist("AssignedTo", "Assigned To", ref addedAttributes);
                    createFieldString("MainContactPhone_1", "Main Contact Phone 1", StringFormat.Phone, ref addedAttributes);
                    createFieldString("MainContactPhone_2", "Main Contact Phone 2", StringFormat.Phone, ref addedAttributes);
                    createFieldDate("UploadedDate", "Uploaded Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldString("LegalDescription", "Legal Description", StringFormat.Text, ref addedAttributes, "multi");
                    createFieldString("RPROwnerName", "RPR Owner Name", StringFormat.Text, ref addedAttributes);
                    createFieldString("LoanToValue", "Loan To Value", StringFormat.Text, ref addedAttributes);
                    createFieldString("LoanToValue_percent", "Loan To Value(%)", StringFormat.Text, ref addedAttributes);
                    createFieldString("RPRLoan1Amount", "RPR Loan 1 Amount", StringFormat.Text, ref addedAttributes, "money");
                    createFieldPicklist("NeedGenealogySearch", "Need Genealogy Search?", ref addedAttributes);
                    createFieldString("Notefortheteam", "Note for the team", StringFormat.Text, ref addedAttributes, "multi");
                    createFieldString("OutstandingLoan2", "Outstanding Loan 2", StringFormat.Text, ref addedAttributes, "money");
                    createFieldDate("RecordedLoan1Date", "Recorded Loan 1 Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldDate("RecordedLoan2Date", "Recorded Loan 2 Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldString("LendersName2", "Lender's Name 2", StringFormat.Text, ref addedAttributes);
                    createFieldString("OutstandingLoan1", "Outstanding Loan 1", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("RPRLoan2Amount", "RPR Loan 2 Amount", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("RPRLender2", "RPR Lender 2", StringFormat.Text, ref addedAttributes);
                    createFieldString("LendersName1", "Lender's Name 1", StringFormat.Text, ref addedAttributes);
                    createFieldDate("RPRRecordedLoan2Date", "RPR Recorded Loan 2 Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldDate("RPRRecordedLoan1Date", "RPR Recorded Loan 1 Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldString("RPRLender1", "RPR Lender 1", StringFormat.Text, ref addedAttributes);
                    createFieldString("RealtytracEstimatedProfit", "Realtytrac Estimated Profit", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("RPRLoantoValue_percentage", "RPR Loan to Value (%)", StringFormat.Text, ref addedAttributes);
                    createFieldString("RPREstimatedProfit", "RPR Estimated Profit", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("TotalLoanAmount", "Total Loan Amount", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("TotalLoansAmount", "Total Loans Amount", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("RealtytracEstimatedProfit", "Realtytrac Estimated Profit", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("OutstandingLoan3", "Outstanding Loan 3", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("LendersName3", "Lender's Name 3", StringFormat.Text, ref addedAttributes);
                    createFieldDate("RecordedLoan3Date", "Recorded Loan 3 Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldPicklist("GreatSpread", "Great Spread?", ref addedAttributes);
                    createFieldPicklist("Interested", "Interested?", ref addedAttributes);
                    createFieldDate("DateforGenealogyInfosurplus", "Date Requested for Additional Genealogy Info", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldDate("DateforGenealogyInfo", "Date Requested for Genealogy Info", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldPicklist("Researcher", "Researcher", ref addedAttributes);
                    createFieldString("Equity", "Equity", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("Equitys", "Equitys", StringFormat.Text, ref addedAttributes, "money");
                    createFieldDate("GenealogySearchDone", "Genealogy Search Done", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldPicklist("AdditionalGenealogySearchSource", "Additional Genealogy Search Source", ref addedAttributes);
                    createFieldPicklist("GenealogySearchSources", "Genealogy Search Sources", ref addedAttributes);
                    createFieldDate("AuctionDate", "Auction Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldString("BrokerName", "Broker Name", StringFormat.Text, ref addedAttributes);
                    createFieldDate("SentReferralsAgreement", "Sent Referrals Agreement", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldDate("AgentFoundDate", "Agent Found Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldDate("InitialHOContacted", "Initial HO Contacted", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldDate("ReceivedSignedRefAgreement", "Received Signed Ref Agreement", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldString("ReferralFeeAmount", "Referral Fee Amount", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("BrokerCity", "Broker City", StringFormat.Text, ref addedAttributes);
                    createFieldString("BrokerState", "Broker State", StringFormat.Text, ref addedAttributes);
                    createFieldString("BrokerCellphone", "Broker Cellphone", StringFormat.Phone, ref addedAttributes);
                    createFieldString("BrokerZipcode", "Broker Zipcode", StringFormat.Text, ref addedAttributes);
                    createFieldString("BrokerAddress", "Broker Address", StringFormat.Text, ref addedAttributes);
                    createFieldString("CallerNotes", "Caller Notes", StringFormat.Text, ref addedAttributes, "multi");
                    createFieldDate("LastSoldDate", "Last Sold Date", ref addedAttributes, DateTimeFormat.DateOnly);
                    createFieldString("LastSoldPrice", "Last Sold Price", StringFormat.Text, ref addedAttributes, "money");
                    createFieldString("NumberofBeds", "Number of Beds", StringFormat.Text, ref addedAttributes);
                    createFieldString("NumberofBaths", "Number of Baths", StringFormat.Text, ref addedAttributes);
                    // @todo Next action needed

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

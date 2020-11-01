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
                        IsGlobal = true,
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
                        IsGlobal = true,
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
                    createFieldString("ContactID", "Contact ID", StringFormat.Text, ref addedAttributes);
                    createFieldString("ContactOwnerID", "Contact Owner ID", StringFormat.Text, ref addedAttributes);
                    createFieldPicklist("LeadSource", "Lead Source", "new_msvproperties_leadtype_option_global20201024084509653", ref addedAttributes );
                    createFieldString("Email", "Email", StringFormat.Email, ref addedAttributes );
                    createFieldString("Phone", "Phone", StringFormat.Phone, ref addedAttributes );
                    createFieldString("Mobile", "Mobile", StringFormat.Phone,ref  addedAttributes );
                    createFieldDate("CreatedTime", "Created Time", ref addedAttributes, DateTimeFormat.DateAndTime);
					createFieldString("FullName", "Full Name", StringFormat.Text, ref addedAttributes );
					createFieldString("MailingStreet", "Mailing Street", StringFormat.Text, ref addedAttributes );
					createFieldString("Description", "Description", StringFormat.Text, ref addedAttributes, "multi" );
					createFieldString("Email", "Email", StringFormat.Email, ref addedAttributes );
					createFieldString("SecondaryEmail", "Secondary Email", StringFormat.Email, ref addedAttributes );
					createFieldString("ZillowPropertyID", "Zillow Property ID", StringFormat.Text, ref addedAttributes );
					createFieldString("ZillowHomeValue", "Zillow Home Value", StringFormat.Text, ref addedAttributes, "money" );
					createFieldString("ZillowValueRangeLow", "Zillow Value Range Low", StringFormat.Text, ref addedAttributes, "money" );
					createFieldString("ZillowBathroom", "Zillow Bathroom", StringFormat.Text, ref addedAttributes );
					createFieldString("ZillowBedroom", "Zillow Bedroom", StringFormat.Text, ref addedAttributes );
					createFieldDate("ZillowLastSoldDate", "Zillow Last Sold Date", ref addedAttributes, DateTimeFormat.DateOnly);
					createFieldString("ZillowLastSoldPrice", "Zillow Last Sold Price", StringFormat.Text, ref addedAttributes, "money" );
					createFieldString("ZillowLotSize", "Zillow Lot Size", StringFormat.Text, ref addedAttributes );
					createFieldString("ZillowTaxAssessment", "Zillow Tax Assessment", StringFormat.Text, ref addedAttributes, "money" );
					createFieldDate("ZillowTaxAssessment", "Zillow Tax Assessment", ref addedAttributes, DateTimeFormat.DateOnly);
					createFieldString("ZillowValueRangeHigh", "Zillow Value Range High", StringFormat.Text, ref addedAttributes, "money" );
					createFieldString("ZillowYearBuilt", "Zillow Year Built", StringFormat.Text, ref addedAttributes);
					createFieldString("ZillowLink", "Zillow Link", StringFormat.Url, ref addedAttributes );
					createFieldString("DocketNumber", "Docket Number", StringFormat.Text, ref addedAttributes);
					createFieldDate("DateofDeath", "Date of Death", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("ProbateDate", "Probate Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("ProbateCounty", "Probate County", StringFormat.Text, ref addedAttributes);
					createFieldString("LeadSourceID", "Lead Source ID", StringFormat.Text, ref addedAttributes);
					createFieldString("LeadNumber", "Lead Number", StringFormat.Text, ref addedAttributes);
					createFieldPicklist("SalesStage", "Sales Stage", "cr4f2_opportunitysalesstage", ref addedAttributes);
					createFieldString("TextResponse", "Text Response", StringFormat.Text, ref addedAttributes, "multi");
					createFieldPicklist("OpportunityStatus", "Status", "new_msvproperties_status_option_global", ref addedAttributes);
					createFieldString("DealsID", "Deals ID", StringFormat.Text, ref addedAttributes);
					createFieldString("AttorneyCityTown", "Attorney City/Town", StringFormat.Text, ref addedAttributes);
					createFieldString("AttorneyAddress", "Attorney Address", StringFormat.Text, ref addedAttributes);
					createFieldString("AttorneyZipCode", "Attorney Zip Code", StringFormat.Text, ref addedAttributes);
					createFieldString("AttorneyStateProvince", "Attorney State/Province", StringFormat.Text, ref addedAttributes);
					createFieldString("Zestimate", "Zestimate", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("AttorneyFirmName", "Attorney Firm Name", StringFormat.Text, ref addedAttributes);
					createFieldString("AttorneyPhoneNumber", "Attorney Phone Number", StringFormat.Phone, ref addedAttributes);
					createFieldString("AttorneyPrimaryEmail", "Attorney Primary Email", StringFormat.Email, ref addedAttributes);
					createFieldString("AttorneyFullName", "Attorney Full Name", StringFormat.Text, ref addedAttributes);
					createFieldString("DecendantName", "Decendant Name", StringFormat.Text, ref addedAttributes);
					createFieldString("FirstBankPayoff", "1st Bank Payoff", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("firstbanklender", "1st Bank/Lender", StringFormat.Text, ref addedAttributes);
					createFieldDate("AuctionDate", "Auction Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("PostponedUntil", "Postponed Until", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldPicklist("assignedTo", "Assigned To", "assignedTo", ref addedAttributes);
					createFieldString("Phone_2", "Phone_2", StringFormat.Phone, ref addedAttributes);
					createFieldString("Mobile_2", "Mobile_2", StringFormat.Phone, ref addedAttributes);
					createFieldString("PropertySnapshot", "Property Snapshot", StringFormat.Text, ref addedAttributes, "multi");
					createFieldString("LegalDescription", "Legal Description", StringFormat.Text, ref addedAttributes, "multi");
					createFieldString("secondBankPayoff", "2nd Bank Payoff", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("secondBankLender", "2nd Bank / Lender", StringFormat.Text, ref addedAttributes);
					createFieldString("County", "County", StringFormat.Text, ref addedAttributes);
					createFieldPicklist("Neighborhood", "Neighborhood", "new_msvproperties_neighborhood_option_global", ref addedAttributes);
					createFieldString("NumberContacted", "Number Contacted", StringFormat.Phone, ref addedAttributes);
					createFieldString("BrokerOfficeNumber", "Broker Office Number", StringFormat.Phone, ref addedAttributes);
					createFieldString("LockboxCode", "Lockbox Code", StringFormat.Text, ref addedAttributes);
					createFieldString("DueDilligencePeriod", "Due Dilligence Period", StringFormat.Text, ref addedAttributes);
					createFieldString("FullContractPrice", "Full Contract Price", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("DefaultClosingDays", "Default Closing Days", StringFormat.Text, ref addedAttributes);
					createFieldString("titlecompanynameid", "Title Company Name ID", StringFormat.Text, ref addedAttributes);
					createFieldString("Balanceatclosing", "Balance at Closing", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("Deposit", "Deposit", StringFormat.Text, ref addedAttributes, "money");
					createFieldPicklist("allowedbusinessStates", "Title Company is Licensed to do business in", "new_labelsusastates", ref addedAttributes );
					createFieldString("TitleState", "Title State", StringFormat.Text, ref addedAttributes);
					createFieldString("TitleZipCode", "Title Zip Code", StringFormat.Text, ref addedAttributes);
					createFieldString("TitleStreet", "Title Street", StringFormat.Text, ref addedAttributes);
					createFieldString("titleCity", "Title City", StringFormat.Text, ref addedAttributes);
					createFieldString("FullName", "FullName", StringFormat.Text, ref addedAttributes);
					createFieldDate("ClosingDateManual", "Closing Date(Manual)", ref addedAttributes, DateTimeFormat.DateOnly);
					createFieldString("ZillowRentalHomeValue", "Zillow Rental Home Value", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("ZillowrentalValueRangeHigh", "Zillow Rental Value Range High", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("ZillowrentalValueRangeLow", "Zillow Rental Value Range Low", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("Outstanding", "Outstanding", StringFormat.Text, ref addedAttributes, "money");
					createFieldDate("ContractSent", "Contract Sent", ref addedAttributes, DateTimeFormat.DateOnly);
					createFieldString("1stloanNumber", "1st Loan Number", StringFormat.Text, ref addedAttributes);
					createFieldPicklist("Referralpercentagepick", "Referral", "new_referralmsvprop", ref addedAttributes);
					createFieldDate("ClosingDate", "Closing Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("AVMRVM", "AVM/RVM", StringFormat.Text, ref addedAttributes);
					createFieldString("2ndOwnersName", "2nd Owner's Name", StringFormat.Text, ref addedAttributes);
					createFieldString("3rdOwnersName", "3rd Owner's Name", StringFormat.Text, ref addedAttributes);
					createFieldString("4rthOwnersName", "4rth Owner's Name", StringFormat.Text, ref addedAttributes);
					createFieldBoolean("NeedsAnEval", "Needs An Eval?", "Yes", "No", ref addedAttributes);
					createFieldPicklist("ConsiderPurchase", "Consider Purchase Checklist", "new_considerpurchasechecklistglobal", ref addedAttributes, "multi");
					createFieldDate("ClosingDateAuto", "Closing Date(Auto)", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("ADContractPrice", "AD Contract Price", StringFormat.Text, ref addedAttributes, "money");
					createFieldBoolean("HotLead", "Hot Lead?", "Yes", "No", ref addedAttributes);
					createFieldBoolean("postponementneed", "Needs Postponement?", "Yes", "No", ref addedAttributes);
					createFieldBoolean("Get", "Get?", "Yes", "No", ref addedAttributes);
					createFieldBoolean("JudicialStateYN", "Judicial State(Y/N)", "Yes", "No", ref addedAttributes);
					createFieldDate("ReceivedSignedContract", "Received Signed Contract", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("RPROwnerNames", "RPR Owner Name/s", StringFormat.Text, ref addedAttributes);
					createFieldString("Latest2Notes", "Latest 2 Notes", StringFormat.Text, ref addedAttributes, "multi");
					createFieldString("Balanceatclosing", "Balance At Closing 2", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("Balanceatclosingauto", "Balance At Closing Auto", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("SpecialTermsofContract", "Special Terms of Contract", StringFormat.Text, ref addedAttributes, "multi");
					createFieldDate("ContractRequested", "Contract Requested", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("RealtorcomEstimate", "Realtor.com Estimate", StringFormat.Text, ref addedAttributes);
					createFieldString("ObitLink", "Obit Link", StringFormat.Url, ref addedAttributes);
					createFieldString("Age", "Age", StringFormat.Text, ref addedAttributes);
					createFieldBoolean("BankruptYesNo", "Bankrupt (Yes/No)", "Yes", "No", ref addedAttributes );
					createFieldDate("Sent3PA", "Sent 3PA", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("Sent3PAL2", "Sent 3PA (L2)", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("Authorized3PA", "Authorized 3PA", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("Authorized3PAL2", "Authorized 3PA(L2)", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("DateUploaded", "Date Uploaded", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("2ndLoanNumber", "2nd Loan Number", StringFormat.Text, ref addedAttributes);
					createFieldDate("ReceivedPayoffReinstatement", "Received Payoff/Reinstatement", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("RequestedPayoffReinstatement", "Requested Payoff/Reinstatement", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("ReceivedPayoffReinstatement2", "Received Payoff/Reinstatement2", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("RequestedPayoffReinstatement2", "Requested Payoff/Reinstatement2", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("LenderEmail", "Lender Email", StringFormat.Email, ref addedAttributes);
					createFieldString("LenderEmail2", "Lender 2 Email", StringFormat.Email, ref addedAttributes);
					createFieldString("FaxForeclosureDept", "Fax(Foreclosure Dept)", StringFormat.Phone, ref addedAttributes);
					createFieldString("BankLenderPhone1", "Bank/Lender Phone1", StringFormat.Phone, ref addedAttributes);
					createFieldString("ContactPerson", "Contact Person", StringFormat.Text, ref addedAttributes);
					createFieldString("ContactPerson2", "Contact Person 2", StringFormat.Text, ref addedAttributes);
					createFieldString("FaxNoLender1", "Fax No./Lender 1", StringFormat.Phone, ref addedAttributes);
					createFieldString("FaxNo", "Fax No.", StringFormat.Phone, ref addedAttributes);
					createFieldString("FaxForeclosureDeptLender1", "Fax (Foreclosure Dept.) / Lender 1", StringFormat.Phone, ref addedAttributes);
					createFieldString("Lender2Notes", "Lender 2 Notes", StringFormat.Text, ref addedAttributes, "multi");
					createFieldString("Notes", "Notes", StringFormat.Text, ref addedAttributes, "multi");
					createFieldString("PropStreamValue", "PropStream Value", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("1stBankLenderPhone2", "1st Bank/Lender Phone 2", StringFormat.Phone, ref addedAttributes);
					createFieldString("BankLenderPhone2", "Bank/Lender Phone 2", StringFormat.Phone, ref addedAttributes);
					createFieldString("BankLenderID", "Bank/Lender ID", StringFormat.Text, ref addedAttributes);
					createFieldString("FaxNumber", "Fax Number", StringFormat.Phone, ref addedAttributes);
					createFieldString("FaxForeclosureDept", "Fax(Foreclosure Dept.)", StringFormat.Phone, ref addedAttributes);
					createFieldString("BankLenderPhone1", "Bank / LenderPhone1", StringFormat.Phone, ref addedAttributes);
					createFieldDate("OrderedListingDate", "Ordered Listing Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("MLSSite", "MLS Site", StringFormat.Text, ref addedAttributes);
					createFieldDate("ApprovedtoListDate", "Approved to List Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("ListedDate", "Listed Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldPicklist("ListingStatus", "Listing Status", "new_listingstatus", ref addedAttributes );
					createFieldString("ListedPrice", "Listed Price", StringFormat.Text, ref addedAttributes, "money");
					createFieldDate("PayoffGoodThru", "Payoff Good Thru", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("PayoffGoodThru2", "Payoff Good Thru 2", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("LockboxInstalledDate", "Lockbox Installed Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("SoldDate", "Sold Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldString("BrokerName", "Broker Name", StringFormat.Text, ref addedAttributes );
					createFieldPicklist("ReferralsAgreement", "Referrals Agreement", "new_referralmsvprop", ref addedAttributes);
					createFieldPicklist("ReasonsofContractSigningDelay", "Reasons of Contract Signing Delay", "new_reasonsdelayforsigningsofcontract", ref addedAttributes);
					createFieldDate("InitialHOContacted", "Initial HO Contacted", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("LastHOContacted", "Last HO Contacted", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("SentReferralsAgreement", "Sent Referrals Agreement", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("AgentFoundDate", "Agent Found Date", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("ScheduledMeeting", "Scheduled Meeting", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldDate("ReceivedSignedRefAgreement", "Received Signed Ref Agreement", ref addedAttributes, DateTimeFormat.DateOnly );
					createFieldPicklist("RefertoBrokerStatus", "Refer to Broker Status", "new_refertobrokerstatus", ref addedAttributes);
					createFieldPicklist("LeadDataSource", "Lead Data Source", "new_leadsource_option_132480270695324980", ref addedAttributes);
					createFieldString("LatestNotes", "LatestNotes", StringFormat.Text, ref addedAttributes, "multi");
					createFieldBoolean("VacantOrOccupied", "Vacant or Occupied?", "Vacant", "Occupied", ref addedAttributes);
					createFieldPicklist("LeadType", "Lead Type", "new_msvproperties_leadtype_option_global20201024084509653", ref addedAttributes);
					createFieldString("ApproxEquity", "Approx Equity", StringFormat.Text, ref addedAttributes, "money");
					createFieldString("BrokerAddress", "Broker Address", StringFormat.Text, ref addedAttributes);
					createFieldString("BrokerCity", "Broker City", StringFormat.Text, ref addedAttributes);
					createFieldString("BrokerState", "XAX", StringFormat.Text, ref addedAttributes);
					createFieldString("BrokerZipcode", "Broker Zipcode", StringFormat.Text, ref addedAttributes);
					createFieldString("Longitude", "Longitude", StringFormat.Text, ref addedAttributes);
					createFieldString("Latitude", "Latitude", StringFormat.Text, ref addedAttributes);
					createFieldString("Distance", "Distance", StringFormat.Text, ref addedAttributes);
					createFieldBoolean("CalculateDistanceofAgents", "Calculate Distance of Agents", "Yes", "No", ref addedAttributes);
					createFieldString("SellerName1", "Seller Name 1", StringFormat.Text, ref addedAttributes);
					createFieldString("SellerLink1", "Seller Link 1", StringFormat.Url, ref addedAttributes);
					createFieldString("SellerName2", "Seller Name 2", StringFormat.Text, ref addedAttributes);
					createFieldString("SellerLink2", "Seller Link 2", StringFormat.Url, ref addedAttributes);
					createFieldString("SellerName3", "Seller Name 3", StringFormat.Text, ref addedAttributes);
					createFieldString("SellerLink3", "Seller Link 3", StringFormat.Url, ref addedAttributes);
					createFieldString("SellerName4", "Seller Name 4", StringFormat.Text, ref addedAttributes);
					createFieldString("SellerLink4", "Seller Link 4", StringFormat.Url, ref addedAttributes);
					createFieldString("ContractCreatedBy", "Contract Created By", StringFormat.Email, ref addedAttributes);
					createFieldString("ContractLink", "Contract Link", StringFormat.Url, ref addedAttributes);
					createFieldPicklist("ContractStatus", "Contract Status", "new_contractstatus", ref addedAttributes);
                    #endregion FieldImport

                    List<string> attributesnotAdded = new List<string>();
                    int attributesCount = addedAttributes.Count;
                    int i = 1; // Counter.
                    foreach (AttributeMetadata addedAttributesTuples in addedAttributes)
                    {
                        // Create the request.
                        CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                        {
                            EntityName = Lead.EntityLogicalName,
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
                    Console.WriteLine("Published all customizations for MSVproperties for the <{0}> entity.", Lead.EntityLogicalName);

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

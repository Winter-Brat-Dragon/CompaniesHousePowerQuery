let
    // ---------------------------
    // PARAMETERS
    // ---------------------------
    InputTableName = "CompanyTable",

    // ---------------------------
    // LOAD INPUT TABLE
    // ---------------------------
    Source = Excel.CurrentWorkbook(){[Name=InputTableName]}[Content],
    ChangeType = Table.TransformColumnTypes(Source, {{"CompanyNumber", type text}}),

    // ---------------------------
    // CALL OFFICERS API AND HANDLE ERRORS
    // ---------------------------
    AddOfficersData = Table.AddColumn(
        ChangeType,
        "OfficersData",
        each try GetCompanyOfficers([API_URL_Officers]) otherwise {}
    ),

    // ---------------------------
    // REMOVE RAW INPUT COLUMNS
    // ---------------------------
    RemoveRawInputColumns = Table.RemoveColumns(
        AddOfficersData,
        {"CompanyName", "API_URL", "API_URL_Officers", "API_URL_PSC"}
    ),

    // ---------------------------
    // EXPAND OFFICERS LIST
    // ---------------------------
    ExpandOfficers = Table.ExpandListColumn(RemoveRawInputColumns, "OfficersData"),

    // ---------------------------
    // MAPPING TABLE FOR OFFICERS
    // ---------------------------
    MappingTable = Table.FromRecords({
        [Column="OfficerName", Path={"name"}, Default=""],
        [Column="OfficerRole", Path={"officer_role"}, Default=""],
        [Column="DateOfBirth_Month", Path={"date_of_birth","month"}, Default=null],
        [Column="DateOfBirth_Year", Path={"date_of_birth","year"}, Default=null],
        [Column="AppointedOn", Path={"appointed_on"}, Default=null],
        [Column="ResignedOn", Path={"resigned_on"}, Default=null],
        [Column="Nationality", Path={"nationality"}, Default=""],
        [Column="CountryOfResidence", Path={"country_of_residence"}, Default=""],
        [Column="AppointmentVerificationStartOn", Path={"identity_verification_details","appointment_verification_start_on"}, Default=null],
        [Column="AppointmentVerificationStatementDueOn", Path={"identity_verification_details","appointment_verification_statement_due_on"}, Default=null],
        [Column="AuthorisedCorporateServiceProviderName", Path={"identity_verification_details","authorised_corporate_service_provider_name"}, Default=""],
        [Column="IdentityVerifiedOn", Path={"identity_verification_details","identity_verified_on"}, Default=null],
        [Column="PreferredName", Path={"identity_verification_details","preferred_name"}, Default=""],
        [Column="IdentityType", Path={"identification","identification_type"}, Default=""],
        [Column="IdentificationNumber", Path={"identification","registration_number"}, Default=""],
        [Column="Address_Line1", Path={"address","address_line_1"}, Default=""],
        [Column="Address_Line2", Path={"address","address_line_2"}, Default=""],
        [Column="Address_Town", Path={"address","locality"}, Default=""],
        [Column="Address_Country", Path={"address","country"}, Default=""],
        [Column="Address_PostCode", Path={"address","postal_code"}, Default=""]
    }),

    // ---------------------------
    // ADD MAPPED OFFICER COLUMNS
    // ---------------------------
    AddMappedOfficerColumns = List.Accumulate(
        Table.ToRecords(MappingTable),
        ExpandOfficers,
        (state, mapping) =>
            Table.AddColumn(
                state,
                mapping[Column],
                each GetNestedField([OfficersData], mapping[Path], mapping[Default])
            )
    ),

    // ---------------------------
    // DERIVE LOGICAL IDENTITY VERIFIED
    // ---------------------------
    AddIdentityVerified = Table.AddColumn(
        AddMappedOfficerColumns,
        "IdentityVerified",
        each ([AppointmentVerificationStartOn] <> null),
        type logical
    ),

    // ---------------------------
    // REORDER COLUMNS
    // ---------------------------
    ReorderColumns = List.InsertRange(
        List.RemoveItems(Table.ColumnNames(AddIdentityVerified), {"IdentityVerified"}),
        List.PositionOf(Table.ColumnNames(AddIdentityVerified),"CountryOfResidence")+1,
        {"IdentityVerified"}
    ),
    ReorderedTable = Table.ReorderColumns(AddIdentityVerified, ReorderColumns),

    // ---------------------------
    // REMOVE RAW DATA COLUMN
    // ---------------------------
    RemoveRawDataColumn = Table.RemoveColumns(ReorderedTable, {"OfficersData"}),

    // ---------------------------
    // APPLY STRICT DATA TYPES
    // ---------------------------
    ApplyDataTypes = Table.TransformColumnTypes(
        RemoveRawDataColumn,
        {
            {"CompanyNumber", type text},
            {"OfficerName", type text},
            {"OfficerRole", type text},
            {"DateOfBirth_Month", Int64.Type},
            {"DateOfBirth_Year", Int64.Type},
            {"AppointedOn", type date},
            {"ResignedOn", type date},
            {"Nationality", type text},
            {"CountryOfResidence", type text},
            {"IdentityVerified", type logical},
            {"AppointmentVerificationStartOn", type date},
            {"AppointmentVerificationStatementDueOn", type date},
            {"AuthorisedCorporateServiceProviderName", type text},
            {"IdentityVerifiedOn", type date},
            {"PreferredName", type text},
            {"IdentityType", type text},
            {"IdentificationNumber", type text},
            {"Address_Line1", type text},
            {"Address_Line2", type text},
            {"Address_Town", type text},
            {"Address_Country", type text},
            {"Address_PostCode", type text}
        }
    ),

    // ---------------------------
    // SORT ACTIVE OFFICERS FIRST, THEN ALPHABETICALLY
    // ---------------------------
    SortedTable = Table.Sort(
        ApplyDataTypes,
        {
            {"ResignedOn", Order.Ascending}, // null = active first
            {"OfficerName", Order.Ascending}
        }
    )

in
    SortedTable

    /*

Column Reference Table: CompanyOfficers

| Column Name                          | Source Path / Description                                           |
|--------------------------------------|--------------------------------------------------------------------|
| CompanyNumber                         | Input Table (CompanyTable)                                         |
| OfficerName                           | OfficersData.name                                                  |
| OfficerRole                           | OfficersData.officer_role                                          |
| DateOfBirth_Month                      | OfficersData.date_of_birth.month                                    |
| DateOfBirth_Year                       | OfficersData.date_of_birth.year                                     |
| AppointedOn                            | OfficersData.appointed_on                                          |
| ResignedOn                             | OfficersData.resigned_on                                           |
| Nationality                            | OfficersData.nationality                                           |
| CountryOfResidence                     | OfficersData.country_of_residence                                  |
| IdentityVerified                        | Derived: AppointmentVerificationStartOn <> null                   |
| AppointmentVerificationStartOn         | OfficersData.identity_verification_details.appointment_verification_start_on |
| AppointmentVerificationStatementDueOn  | OfficersData.identity_verification_details.appointment_verification_statement_due_on |
| AuthorisedCorporateServiceProviderName | OfficersData.identity_verification_details.authorised_corporate_service_provider_name |
| IdentityVerifiedOn                     | OfficersData.identity_verification_details.identity_verified_on    |
| PreferredName                          | OfficersData.identity_verification_details.preferred_name          |
| IdentityType                           | OfficersData.identification.identification_type                   |
| IdentificationNumber                   | OfficersData.identification.registration_number                   |
| Address_Line1                           | OfficersData.address.address_line_1                                 |
| Address_Line2                           | OfficersData.address.address_line_2                                 |
| Address_Town                            | OfficersData.address.locality                                       |
| Address_Country                         | OfficersData.address.country                                        |
| Address_PostCode                        | OfficersData.address.postal_code                                     |

*/

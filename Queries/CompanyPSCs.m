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
    // CALL PSC API AND HANDLE ERRORS
    // ---------------------------
    AddPSCsData = Table.AddColumn(
        ChangeType,
        "PSCsData",
        each try GetCompanyPSCs([API_URL_PSC]) otherwise {}
    ),

    // ---------------------------
    // REMOVE RAW INPUT COLUMNS
    // ---------------------------
    RemoveRawInputColumns = Table.RemoveColumns(
        AddPSCsData,
        {"CompanyName", "API_URL", "API_URL_Officers", "API_URL_PSC"}
    ),

    // ---------------------------
    // EXPAND PSC LIST
    // ---------------------------
    ExpandPSCs = Table.ExpandListColumn(RemoveRawInputColumns, "PSCsData"),

    // ---------------------------
    // MAPPING TABLE FOR PSC FIELDS
    // ---------------------------
    MappingTable = Table.FromRecords({
        [Column="Name", Path={"name"}, Default=""],
        [Column="Title", Path={"name_elements","title"}, Default=""],
        [Column="Forename", Path={"name_elements","forename"}, Default=""],
        [Column="MiddleName", Path={"name_elements","middle_name"}, Default=""],
        [Column="Surname", Path={"name_elements","surname"}, Default=""],
        [Column="NotifiedOn", Path={"notified_on"}, Default=null],
        [Column="Ceased", Path={"ceased"}, Default=false],
        [Column="Ceased_On", Path={"ceased_on"}, Default=null],
        [Column="DateOfBirth_Month", Path={"date_of_birth","month"}, Default=null],
        [Column="DateOfBirth_Year", Path={"date_of_birth","year"}, Default=null],
        [Column="Nationality", Path={"nationality"}, Default=""],
        [Column="CountryOfResidence", Path={"country_of_residence"}, Default=""],
        [Column="AppointmentVerificationStartOn", Path={"identity_verification_details","appointment_verification_start_on"}, Default=null],
        [Column="AppointmentVerificationStatementDate", Path={"identity_verification_details","appointment_verification_statement_date"}, Default=null],
        [Column="AppointmentVerificationStatementDueOn", Path={"identity_verification_details","appointment_verification_statement_due_on"}, Default=null],
        [Column="AuthorisedCorporateServiceProviderName", Path={"identity_verification_details","authorised_corporate_service_provider_name"}, Default=""],
        [Column="IdentityVerifiedOn", Path={"identity_verification_details","identity_verified_on"}, Default=null],
        [Column="PreferredName", Path={"identity_verification_details","preferred_name"}, Default=""],
        [Column="IdentificationType", Path={"identification","identification_type"}, Default=""],
        [Column="IdentificationRegistrationNumber", Path={"identification","registration_number"}, Default=""],
        [Column="Kind", Path={"kind"}, Default=""],
        [Column="NaturesOfControl", Path={"natures_of_control"}, Default={} ],
        [Column="Address_Line1", Path={"address","address_line_1"}, Default=""],
        [Column="Address_Line2", Path={"address","address_line_2"}, Default=""],
        [Column="Address_Town", Path={"address","locality"}, Default=""],
        [Column="Address_Country", Path={"address","country"}, Default=""],
        [Column="Address_PostCode", Path={"address","postal_code"}, Default=""],
        [Column="IsSanctioned", Path={"is_sanctioned"}, Default=false]
    }),

    // ---------------------------
    // ADD MAPPED PSC COLUMNS
    // ---------------------------
    AddMappedPSCsColumns = List.Accumulate(
        Table.ToRecords(MappingTable),
        ExpandPSCs,
        (state, mapping) =>
            Table.AddColumn(
                state,
                mapping[Column],
                each if mapping[Column] = "NaturesOfControl" then
                        Text.Combine(GetNestedField([PSCsData], mapping[Path], {}), ", ")
                     else
                        GetNestedField([PSCsData], mapping[Path], mapping[Default])
            )
    ),

    // ---------------------------
    // DERIVE LOGICAL IDENTITY VERIFIED
    // ---------------------------
    AddIdentityVerified = Table.AddColumn(
        AddMappedPSCsColumns,
        "IdentityVerified",
        each ([AppointmentVerificationStartOn] <> null),
        type logical
    ),

    // ---------------------------
    // REORDER COLUMNS
    // ---------------------------
    ReorderColumns = {
        "CompanyNumber",
        "Name","Title","Forename","MiddleName","Surname",
        "NotifiedOn","Ceased","Ceased_On",
        "DateOfBirth_Month","DateOfBirth_Year",
        "Nationality","CountryOfResidence","IdentityVerified",
        "AppointmentVerificationStartOn","AppointmentVerificationStatementDate","AppointmentVerificationStatementDueOn",
        "AuthorisedCorporateServiceProviderName","IdentityVerifiedOn","PreferredName",
        "IdentificationType","IdentificationRegistrationNumber",
        "Kind","NaturesOfControl",
        "Address_Line1","Address_Line2","Address_Town","Address_Country","Address_PostCode",
        "IsSanctioned"
    },
    ReorderedTable = Table.ReorderColumns(AddIdentityVerified, ReorderColumns),

    // ---------------------------
    // REMOVE RAW DATA COLUMN
    // ---------------------------
    RemoveRawDataColumn = Table.RemoveColumns(ReorderedTable, {"PSCsData"}),

    // ---------------------------
    // APPLY STRICT DATA TYPES
    // ---------------------------
    ApplyDataTypes = Table.TransformColumnTypes(
        RemoveRawDataColumn,
        {
            {"CompanyNumber", type text},
            {"Name", type text},
            {"Title", type text},
            {"Forename", type text},
            {"MiddleName", type text},
            {"Surname", type text},
            {"NotifiedOn", type date},
            {"Ceased", type logical},
            {"Ceased_On", type date},
            {"DateOfBirth_Month", Int64.Type},
            {"DateOfBirth_Year", Int64.Type},
            {"Nationality", type text},
            {"CountryOfResidence", type text},
            {"IdentityVerified", type logical},
            {"AppointmentVerificationStartOn", type date},
            {"AppointmentVerificationStatementDate", type date},
            {"AppointmentVerificationStatementDueOn", type date},
            {"AuthorisedCorporateServiceProviderName", type text},
            {"IdentityVerifiedOn", type date},
            {"PreferredName", type text},
            {"IdentificationType", type text},
            {"IdentificationRegistrationNumber", type text},
            {"Kind", type text},
            {"NaturesOfControl", type text},
            {"Address_Line1", type text},
            {"Address_Line2", type text},
            {"Address_Town", type text},
            {"Address_Country", type text},
            {"Address_PostCode", type text},
            {"IsSanctioned", type logical}
        }
    ),

    // ---------------------------
    // SORT ACTIVE PSCS FIRST, THEN ALPHABETICALLY BY Surname
    // ---------------------------
    SortedTable = Table.Sort(
        ApplyDataTypes,
        {
            {"Ceased", Order.Ascending}, // false = active first
            {"Surname", Order.Ascending}
        }
    )

in
    SortedTable

    /*

Column Reference Table: CompanyPSCs

| Column Name                          | Source Path / Description                                           |
|--------------------------------------|--------------------------------------------------------------------|
| CompanyNumber                         | Input Table (CompanyTable)                                         |
| Name                                  | PSCsData.name                                                      |
| Title                                 | PSCsData.name_elements.title                                        |
| Forename                              | PSCsData.name_elements.forename                                      |
| MiddleName                            | PSCsData.name_elements.middle_name                                   |
| Surname                               | PSCsData.name_elements.surname                                       |
| NotifiedOn                            | PSCsData.notified_on                                               |
| Ceased                                | PSCsData.ceased                                                    |
| Ceased_On                             | PSCsData.ceased_on                                                 |
| DateOfBirth_Month                      | PSCsData.date_of_birth.month                                        |
| DateOfBirth_Year                       | PSCsData.date_of_birth.year                                         |
| Nationality                            | PSCsData.nationality                                               |
| CountryOfResidence                     | PSCsData.country_of_residence                                      |
| IdentityVerified                        | Derived: AppointmentVerificationStartOn <> null                   |
| AppointmentVerificationStartOn         | PSCsData.identity_verification_details.appointment_verification_start_on |
| AppointmentVerificationStatementDate   | PSCsData.identity_verification_details.appointment_verification_statement_date |
| AppointmentVerificationStatementDueOn  | PSCsData.identity_verification_details.appointment_verification_statement_due_on |
| AuthorisedCorporateServiceProviderName | PSCsData.identity_verification_details.authorised_corporate_service_provider_name |
| IdentityVerifiedOn                     | PSCsData.identity_verification_details.identity_verified_on        |
| PreferredName                          | PSCsData.identity_verification_details.preferred_name             |
| IdentificationType                     | PSCsData.identification.identification_type                        |
| IdentificationRegistrationNumber       | PSCsData.identification.registration_number                        |
| Kind                                   | PSCsData.kind                                                      |
| NaturesOfControl                        | PSCsData.natures_of_control (concatenated as text)                 |
| Address_Line1                           | PSCsData.address.address_line_1                                     |
| Address_Line2                           | PSCsData.address.address_line_2                                     |
| Address_Town                            | PSCsData.address.locality                                           |
| Address_Country                         | PSCsData.address.country                                            |
| Address_PostCode                        | PSCsData.address.postal_code                                        |
| IsSanctioned                            | PSCsData.is_sanctioned                                             |

*/

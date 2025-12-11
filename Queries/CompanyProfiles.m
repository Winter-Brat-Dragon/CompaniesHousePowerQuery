let
  // ---------------------------
  // PARAMETERS
  // ---------------------------
  InputTableName = "CompanyTable",

  // ---------------------------
  // LOAD INPUT TABLE
  // ---------------------------
  Source = Excel.CurrentWorkbook(){[Name=InputTableName]}[Content],
  InputTableChangedType = Table.TransformColumnTypes(Source, {{"CompanyNumber", type text}}),

  // ---------------------------
  // CALL API AND HANDLE ERRORS
  // ---------------------------
  AddCompanyProfile = Table.AddColumn(
    InputTableChangedType,
    "Data",
    each try GetCompanyProfile([API_URL])
      otherwise [
        accounts = [
          accounting_reference_date = [day=null, month=null],
          last_accounts = [period_end_on=null, type=""],
          next_accounts = [due_on=null, overdue=false, period_end_on=null]
        ],
        can_file = false,
        company_name = "",
        company_status = "",
        confirmation_statement = [
          last_made_up_to = null,
          next_due = null,
          next_made_up_to = null,
          overdue = false
        ],
        date_of_creation = null,
        registered_office_address = [
          address_line_1 = "",
          address_line_2 = "",
          country = "",
          locality = "",
          postal_code = ""
        ],
        type = "",
        subtype = "",
        date_of_cessation = null,
        company_status_detail = ""
      ]
  ),

  // ---------------------------
  // REMOVE UNNECESSARY INPUT COLUMNS
  // ---------------------------
  RemoveRawInputColumns = Table.RemoveColumns(AddCompanyProfile, {"CompanyName", "API_URL", "API_URL_Officers", "API_URL_PSC"}),

  // ---------------------------
  // MAPPING TABLE
  // ---------------------------
  MappingTable = Table.FromRecords({
    [Column="CompanyName", Path={"company_name"}, Default=""],
    [Column="CompanyStatus", Path={"company_status"}, Default=""],
    [Column="StatusDetail", Path={"company_status_detail"}, Default=""],
    [Column="CompanyType", Path={"type"}, Default=""],
    [Column="CompanySubtype", Path={"subtype"}, Default=""],
    [Column="DateOfIncorporation", Path={"date_of_creation"}, Default=null],
    [Column="DateOfCessation", Path={"date_of_cessation"}, Default=null],
    [Column="ARD_Day", Path={"accounts","accounting_reference_date","day"}, Default=null],
    [Column="ARD_Month", Path={"accounts","accounting_reference_date","month"}, Default=null],
    [Column="RegisteredOffice_Line1", Path={"registered_office_address","address_line_1"}, Default=""],
    [Column="RegisteredOffice_Line2", Path={"registered_office_address","address_line_2"}, Default=""],
    [Column="RegisteredOffice_Town", Path={"registered_office_address","locality"}, Default=""],
    [Column="RegisteredOffice_Country", Path={"registered_office_address","country"}, Default=""],
    [Column="RegisteredOffice_PostCode", Path={"registered_office_address","postal_code"}, Default=""],
    [Column="AccountsOverdue", Path={"accounts","next_accounts","overdue"}, Default=false],
    [Column="NextAccountsMadeUpTo", Path={"accounts","next_accounts","period_end_on"}, Default=null],
    [Column="NextAccountsDue", Path={"accounts","next_accounts","due_on"}, Default=null],
    [Column="LastAccountsMadeUpTo", Path={"accounts","last_accounts","period_end_on"}, Default=null],
    [Column="LastAccountsType", Path={"accounts","last_accounts","type"}, Default=""],
    [Column="ConfirmationStatementOverdue", Path={"confirmation_statement","overdue"}, Default=false],
    [Column="NextConfirmationStatementMadeUpTo", Path={"confirmation_statement","next_made_up_to"}, Default=null],
    [Column="NextConfirmationStatementDue", Path={"confirmation_statement","next_due"}, Default=null],
    [Column="LastConfirmationStatementMadeUpTo", Path={"confirmation_statement","last_made_up_to"}, Default=null],
    [Column="CanFile", Path={"can_file"}, Default=false]
  }),

  // ---------------------------
  // ADD MAPPED COLUMNS USING CENTRAL FUNCTION
  // ---------------------------
  AddMappedProfileColumns = List.Accumulate(
    Table.ToRecords(MappingTable),
    RemoveRawInputColumns,
    (state, mapping) =>
      Table.AddColumn(
        state,
        mapping[Column],
        each GetNestedField([Data], mapping[Path], mapping[Default])
      )
  ),

  // ---------------------------
  // REMOVE RAW DATA COLUMN
  // ---------------------------
  RemoveRawDataColumn = Table.RemoveColumns(AddMappedProfileColumns, {"Data"}),

  // ---------------------------
  // APPLY STRICT DATA TYPES
  // ---------------------------
  ApplyDataTypes = Table.TransformColumnTypes(
    RemoveRawDataColumn,
    {
      {"CompanyNumber", type text},
      {"CompanyName", type text},
      {"CompanyStatus", type text},
      {"StatusDetail", type text},
      {"CompanyType", type text},
      {"CompanySubtype", type text},
      {"DateOfIncorporation", type date},
      {"DateOfCessation", type date},
      {"ARD_Day", Int64.Type},
      {"ARD_Month", Int64.Type},
      {"RegisteredOffice_Line1", type text},
      {"RegisteredOffice_Line2", type text},
      {"RegisteredOffice_Town", type text},
      {"RegisteredOffice_Country", type text},
      {"RegisteredOffice_PostCode", type text},
      {"AccountsOverdue", type logical},
      {"NextAccountsMadeUpTo", type date},
      {"NextAccountsDue", type date},
      {"LastAccountsMadeUpTo", type date},
      {"LastAccountsType", type text},
      {"ConfirmationStatementOverdue", type logical},
      {"NextConfirmationStatementMadeUpTo", type date},
      {"NextConfirmationStatementDue", type date},
      {"LastConfirmationStatementMadeUpTo", type date},
      {"CanFile", type logical}
    }
  ),

  // ---------------------------
  // FINAL COLUMN ORDER
  // ---------------------------
  FinalTable = Table.ReorderColumns(ApplyDataTypes, {"CompanyNumber"} & MappingTable[Column]),

  // ---------------------------
  // SORT ALPHABETICALLY BY COMPANY NAME
  // ---------------------------
  SortedTable = Table.Sort(FinalTable, {{"CompanyName", Order.Ascending}})
in
  SortedTable

  /*

Column Reference Table: CompanyProfiles

| Column Name                     | Source Path / Description                                             |
|---------------------------------|-----------------------------------------------------------------------|
| CompanyNumber                    | Input Table (CompanyTable)                                            |
| CompanyName                      | Data.company_name                                                     |
| CompanyStatus                    | Data.company_status                                                   |
| StatusDetail                     | Data.company_status_detail                                            |
| CompanyType                      | Data.type                                                             |
| CompanySubtype                   | Data.subtype                                                          |
| DateOfIncorporation              | Data.date_of_creation                                                 |
| DateOfCessation                  | Data.date_of_cessation                                               |
| ARD_Day                          | Data.accounts.accounting_reference_date.day                           |
| ARD_Month                        | Data.accounts.accounting_reference_date.month                         |
| RegisteredOffice_Line1           | Data.registered_office_address.address_line_1                         |
| RegisteredOffice_Line2           | Data.registered_office_address.address_line_2                         |
| RegisteredOffice_Town            | Data.registered_office_address.locality                               |
| RegisteredOffice_Country         | Data.registered_office_address.country                                 |
| RegisteredOffice_PostCode        | Data.registered_office_address.postal_code                             |
| AccountsOverdue                  | Data.accounts.next_accounts.overdue                                    |
| NextAccountsMadeUpTo             | Data.accounts.next_accounts.period_end_on                              |
| NextAccountsDue                  | Data.accounts.next_accounts.due_on                                     |
| LastAccountsMadeUpTo             | Data.accounts.last_accounts.period_end_on                               |
| LastAccountsType                 | Data.accounts.last_accounts.type                                       |
| ConfirmationStatementOverdue     | Data.confirmation_statement.overdue                                    |
| NextConfirmationStatementMadeUpTo| Data.confirmation_statement.next_made_up_to                             |
| NextConfirmationStatementDue     | Data.confirmation_statement.next_due                                   |
| LastConfirmationStatementMadeUpTo| Data.confirmation_statement.last_made_up_to                             |
| CanFile                          | Data.can_file                                                          |

*/

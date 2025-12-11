// ---------------------------
// GetNestedField Function
// ---------------------------
// Safely extracts a nested field from a record; returns defaultValue if path is invalid or record is null
(GetNestedFieldRecord as nullable record, Path as list, DefaultValue as any) as any =>
let
    // Ensure the record is not null
    SafeRecord = if GetNestedFieldRecord = null then [] else GetNestedFieldRecord,

    // Accumulate through the path list to get the nested value
    Result = List.Accumulate(
                Path,
                SafeRecord,
                (State, Current) =>
                    if State = null then null else Record.FieldOrDefault(State, Current, null)
             )
in
    // Return the default value if the result is null
    if Result = null then DefaultValue else Result

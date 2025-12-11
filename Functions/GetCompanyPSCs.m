// ---------------------------
// GetCompanyPSCs Function
// ---------------------------
// Input: API_URL_PSC (text)
// Output: List of PSC records (empty list if no data or error)
(GetPSCsURL as text) as list =>
let
    // 1. Trim and validate input
    CleanURL = Text.Trim(GetPSCsURL),
    ValidURL = if CleanURL = "" then null else CleanURL,

    // 2. Attempt API request
    Response =
        if ValidURL = null then
            null
        else
            try Web.Contents(ValidURL, [ManualStatusHandling={400,401,403,404,429,500,503}])
            otherwise null,

    // 3. Parse JSON if response exists
    JSON =
        if Response = null then
            null
        else
            try Json.Document(Response) otherwise null,

    // 4. Extract the 'items' list safely
    Items =
        if JSON = null then
            {}
        else
            try JSON[items] otherwise {},

    // 5. Ensure output is always a list
    Result =
        if Items is list then
            Items
        else
            {}
in
    Result

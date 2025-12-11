// ---------------------------
// GetCompanyProfile Function
// ---------------------------
// Input: API_URL (text)
// Output: Record with company profile (null if no data or error)
(GetCompanyURL as text) as nullable record =>
let
    // 1. Trim and validate input
    CleanURL = Text.Trim(GetCompanyURL),
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

    // 4. Ensure output is a record
    Result =
        if JSON is record then
            JSON
        else
            null
in
    Result

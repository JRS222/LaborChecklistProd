$base = "https://webeor.usps.gov"
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# 1. frameset shell (sets NSC_* cookie)
$null = Invoke-WebRequest -Uri "$base/GUI_Login_View.asp" `
        -WebSession $session -UseDefaultCredentials -UseBasicParsing

# 2. THE frame that does the Windows-identity -> app-session binding
$main = Invoke-WebRequest -Uri "$base/GUI_Login_Main.asp" `
        -WebSession $session -UseDefaultCredentials -UseBasicParsing
"After main  : $($main.BaseResponse.ResponseUri.AbsoluteUri)  Bytes: $($main.RawContentLength)"

# 3. does the viewer now accept us?
$view = Invoke-WebRequest -Uri "$base/EORViewer/GUI_EORView_Viewer.asp" `
        -WebSession $session -UseDefaultCredentials -UseBasicParsing
"After viewer: $($view.BaseResponse.ResponseUri.AbsoluteUri)  Bytes: $($view.RawContentLength)"

# 4. show us the site dropdown from the landing page
$opts = [regex]::Matches($main.Content, '<option[^>]*value\s*=\s*["'']([^"'']+)["''][^>]*>([^<]+)</option>', 'IgnoreCase')
"Dropdown options found: $($opts.Count)"
$opts | Select-Object -First 10 | ForEach-Object {
    "  {0}  =>  {1}" -f $_.Groups[1].Value, $_.Groups[2].Value.Trim()
}

# 5. cookies now
$session.Cookies.GetCookies($base) | ForEach-Object { "  $($_.Name) = $($_.Value)" }
That changes the picture — and makes it much easier

If the popup is a fallback and the normal flow is: Windows auth → auto-login → site-selection dropdown, then:

1. We don't need to script a login form at all. Skip checklist item C (auth method).
2. The auto-login almost certainly happens in GUI_Login_Main.asp — that's the frame our PowerShell never fetched. Fetching GUI_Login_View.asp returns the <frameset> shell, but the "you're a real user, here's your session" logic runs inside the frmMain frame's page. Our session got the shell, never the main frame, so it never got authenticated at the app level.
3. The site-selection dropdown is our site-ID mapping. If the landing page has <select> with <option value="19010">PITTSBURGH PA P&DC</option>, we can parse it once per session and resolve any site name → ID without any extra requests.

The focused test (much shorter than the earlier probe)

```powershell
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
```

Reading the output:

After viewer FinalUri Verdict
…/EORViewer/GUI_EORView_Viewer.asp ✅ We're in. One-line fix: fetch GUI_Login_Main.asp before anything else. Run the CSV POST and it'll return a real CSV.
…/GUI_Login_View.asp ❌ App still sees us as unauthenticated. The GUI_Login_Main.asp body will tell us why — most likely it contains a "Sign In" link that redirects to the popup. Paste $main.Content.Substring(0,1000).

The dropdown at step 4 answers checklist item B for free. If it prints entries like 19010 => PITTSBURGH PA P&DC, we have the mapping and don't need -SiteId as a workaround — we can take -Site "PITTSBURGH PA P&DC" and resolve the ID from the dropdown HTML.

The cookies at step 5 matter because if the app session cookie changes name or value after step 2, we know GUI_Login_Main.asp is what issues it — which is exactly the confirmation we need to write Connect-WebEor.

What to send back

Just paste the output of the five blocks above. That's it. No Ctrl+U, no DevTools, no cookie transplant.

From that single output I can write:

· Connect-WebEor — creates a session, hits the three URLs in order, verifies auth
· Resolve-SiteId -Name "…" — parses the dropdown and returns the numeric ID
· Get-WebEorCsv -Session $s -SiteId 19010 -Kind Eor -Date (Get-Date).AddDays(-1) — the working POST wrapped in a function
· The exact lines to replace in Machine Checker.ps1

Still outstanding (do these when you can, they're independent):

· Machine Mapping CSV form — Ctrl+U on the Machine Mapping page, paste the <form name="frmTableCSV"> block (checklist A)
· Your Machine Checker.ps1 — the param block + CSV import section (checklist E)
· Date scope — single day (yesterday) or a range? I've defaulted to single-day until you say otherwise (checklist D)
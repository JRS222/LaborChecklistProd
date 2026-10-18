Found it — it's authentication, not the payload

The smoking gun is this line:

```
FinalUri: https://webeor.usps.gov/GUI_Login_View.asp
```

Your GET landed on the login page, not the viewer. So:

· The $session object has cookies (NSC_xfcfps_443, ASPSESSIONIDCGBDSBRS) — but they belong to an unauthenticated ASP session.
· The POST to GUI_Report_CSV.asp reached the server, but the ASP didn't see an authenticated user, so it rejected with Data Source Unrecongized! (which, given the "Please refresh source page" wording, is really the app's generic "your session doesn't match a viewer page" error).
· The regex found nothing in $get.Content because that content is the login page HTML, not the viewer page.

Everything else is correct. Your browser POST and your PowerShell POST used byte-identical field values — tabCsvSource=CLIENT, tabCsvGroup=+ (space), same SQL, same date. That was the whole diagnostic, and it came back clean. The payload is right. The session is the only missing piece.

Once PowerShell is authenticated, the exact POST you already ran will return a CSV.

---

Three paths forward — pick whichever fits

Path 1: Validate the theory in 60 seconds (cookie transplant)

The user's browser is already authenticated. Prove that the browser's cookie is the only difference:

1. In the authenticated WebEOR browser tab: F12 → Application → Cookies → https://webeor.usps.gov
2. Copy the value of ASPSESSIONIDCGBDSBRS (or whatever ASPSESSIONID* is there)
3. In PowerShell:

```powershell
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.Cookies.Add((New-Object System.Net.Cookie(
    "ASPSESSIONIDCGBDSBRS",
    "<paste-the-browser-value-here>",
    "/",
    "webeor.usps.gov")))

# Re-run the same POST as before
$resp = Invoke-WebRequest -Uri "https://webeor.usps.gov/ReportPage/GUI_Report_CSV.asp" `
                          -Method POST -Body $body -WebSession $session `
                          -UseDefaultCredentials -UseBasicParsing
"Status: $($resp.StatusCode)  CT: $($resp.Headers['Content-Type'])  Bytes: $($resp.RawContentLength)"
$resp.Content.Substring(0, [Math]::Min(300, $resp.Content.Length))
```

If you get CT: text/csv (or application/vnd.ms-excel) and the first line is Site,MType,MNo,…, we're done diagnosing. That confirms cookie-only auth and we just need to automate login. If it still fails, the server has a second check tied to the login event (a session variable the ASP sets at login), and we definitely need to script the login flow.

Don't ship this as the solution — the cookie value expires and is per-user. It's a one-time proof.

Path 2: Script the login (the real fix)

This is what will end up in Machine Checker.ps1. I need one thing from you:

Open https://webeor.usps.gov/GUI_Login_View.asp in the browser (it should auto-redirect if you're logged in; use an InPrivate window to see it fresh) and Ctrl+U on the login page. Paste me:

1. The <form …> tag — specifically its action="…" and method
2. Every <input> inside it — name, type, and (for hidden fields) value
3. Whether the page has a 2FA / MFA / token step after password, or a "remember me" checkbox
4. Whether the login is username/password, Windows auth (does the page say "Sign in with your Windows account"?), or something else (smartcard, OTP, SSO)

With that I can write:

```powershell
function Connect-WebEor {
    param([pscredential]$Credential)
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    # 1. GET login page (may set a pre-auth cookie or return a token)
    # 2. POST username/password to the action URL
    # 3. Verify by GET-ing the viewer page and checking FinalUri stays on /EORViewer/
    return $session
}
```

…and then every CSV fetch reuses that $session.

Path 3: Force Negotiate auth (worth one shot, costs 30 seconds)

-UseDefaultCredentials doesn't always force NTLM/Negotiate in PS 5.1. Try it explicitly:

```powershell
$cred = Get-Credential   # enter your USPS account
$get = Invoke-WebRequest -Uri "https://webeor.usps.gov/EORViewer/GUI_EORView_Viewer.asp" `
        -Credential $cred -Authentication Negotiate -UseBasicParsing
"FinalUri: $($get.BaseResponse.ResponseUri.AbsoluteUri)"
```

If FinalUri stays on /EORViewer/GUI_EORView_Viewer.asp, Windows auth works and the fix is trivial — replace -UseDefaultCredentials with -Credential. If it still bounces to GUI_Login_View.asp, WebEOR uses form login and we go with Path 2.

---

What to send back

Priority order — any one of these unblocks the next step:

1. Path 1 result (cookie transplant → CSV or error) — fastest confirmation
2. Path 3 result (FinalUri after Negotiate) — tells us which auth model we're in
3. Path 2 login form — the actual fix, needed for the final script regardless

Once we know the auth mechanism, the remainder is mechanical: the POST payload you already proved out gets wrapped in a Get-WebEorCsv -Site $SiteId -Kind Eor|Mapping -Date $D function, and Machine Checker.ps1 calls it at the top instead of prompting for a file.

One thing still outstanding from earlier, unrelated to auth — checklist item B: the site name → site_id mapping. Since the browser landed you on a login page, that question is still open. Once you're logged in via the browser, check the authenticated viewer page's URL for a site_id=… parameter, or Ctrl+F the viewer's View Source for gstrTableSQL — it will contain START WITH x.site_id = <id> with the real ID substituted. That's the ID we feed into the START WITH clause, replacing 19010. We'll need it for any site other than 19010.
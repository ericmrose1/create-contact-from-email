CREATE CONTACT FROM EMAIL — OUTLOOK WEB ADD-IN v1.2 (NO INSTALLER)

WHY THIS VERSION EXISTS
-----------------------
Your FortiClient security system quarantined the earlier COM installer because it compiled code and registered a local Outlook component. This version does none of that.

It is ordinary HTML/JavaScript plus an Outlook XML manifest. There is:
- no EXE
- no DLL
- no PowerShell
- no CMD/BAT installer
- no Windows registry change
- no local code compilation
- no Microsoft Graph permission
- no Azure/Entra app registration

WHAT IT DOES
------------
1. You open/select an email in Classic Outlook.
2. Click "Create Contact from Email".
3. It reads the current email and extracts the sender/signature information.
4. You review/edit First, Middle, Last, Company, Title, Email, phones, fax, website and address.
5. Click "Create Outlook Contact File".
6. It downloads a standard .VCF contact file.
7. Open that .VCF in Outlook. Outlook displays the normal contact form.
8. Proof it and click Save & Close.

This deliberately leaves the final save to you.

IMPORTANT ONE-TIME REQUIREMENT
------------------------------
Microsoft Outlook web add-ins must load their HTML from an HTTPS web address. Therefore the included SITE folder must be placed on a simple static HTTPS host.

The easiest personal approach is GitHub Pages. No company-wide Outlook deployment is required. Once the site is hosted, open MAKE-MANIFEST.html, paste the HTTPS site address, and it will create manifest.xml for you.

OUTLOOK INSTALLATION AFTER HOSTING
----------------------------------
1. In Classic Outlook choose File > Info > Manage Add-ins.
2. Your browser opens Outlook's Add-ins page.
3. Choose My add-ins.
4. Scroll to Custom Addins.
5. Choose Add a custom add-in > Add from File.
6. Select the manifest.xml created by MAKE-MANIFEST.html.
7. Accept the prompts.
8. Reopen Outlook if necessary.

Microsoft notes that a manually sideloaded add-in can take time to appear in Classic Outlook because of caching.

NEXT
----
If you want, ChatGPT can walk you through the HTTPS hosting step one screen at a time. You do not need to understand the code.

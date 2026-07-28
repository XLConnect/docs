---
description: The surest way to not spill a secret is to never have it
---

# XLConnect Security

Data security is on everyone's mind these days. If users are considering XLConnect at their company, IT will likely have questions. This page is intended to answer most of these. To ensure the details get across correctly this page can get a bit technical. Normal users can skip over this page as it is intended for IT and security professionals.  &#x20;

Security and data security were a key concern when designing XLConnect. Most financial cloud apps that build on other cloud platforms store passwords and client data on their own servers. XLConnect takes a different approach where sensitive data does not pass over our servers if at all possible.&#x20;

Our security strategy is:&#x20;

&#x20;                                     **"The surest way to not spill a secret is to never have it"**

We take this to the limit, for example we only allow integrated logins Microsoft and Xero, as a username/password would require us to store a password.&#x20;

Running XLConnect in Excel instead of a browser allows more flexibility to store data on the client's laptop directly, bypassing XLConnect servers. This means security concerns like ISO270001 are not applicable as we never see let alone store that data.&#x20;

This also applies to privacy regulations like GDPR as we never have clients' sensitive data or web tokens.

This setup eliminates chapters of risks and has allowed us to pass stringent security reviews as imposed by cloud providers like Xero and become certified.

## External Identity Providers

XLConnect uses so-called 'External Identity Providers' to vouch for your identity to us. When you login to the XLConnect, you can choose to authenticate with either your Microsoft or Xero account. You will be forwarded to their login pages in your own browser, so you can be sure we can't see that.&#x20;

Users can verify this by checking the login address and clicking the lock icon for further details about the certificate sued to secure the connection.&#x20;

<img src="../.gitbook/assets/image (8).png" alt="" data-size="original">

<img src="../.gitbook/assets/image (11).png" alt="" data-size="original">

## Built-in OAuth management

XLConnect has built in authentication for the following cloud systems:&#x20;

* Shopify
* Exact Online
* Xero
* Visma.Net
* Visma Nmbrs
* Hubspot
* Microsoft Dynamics 365 Business Central&#x20;

Because the tokens are only ever on the user's computer, users can only pull their own data associated with their own account.&#x20;

### Handling Access and Refresh Tokens

The authentication tokens and data users pull exist only on their laptop. Logging in to cloud systems with these flows does not grant XLConnect B.V. access to that data as is usually the case with OAuth tokens. XLConnect never sees the tokens and does not have access.&#x20;

### Caveat: OAuth without PKCE

Certain cloud systems like Hubspot or Exact Online do not support the `authentication_code` flow with PKCE, intended for desktop apps and SPA's. These systems use the `authorization_code` flow that requires a client secret to authenticate against their authority. This secret cannot be shipped with the client as it could be reverse engineered from the code.&#x20;

For these systems XLConnect uses a 'server relay' where the authorization code is fetched from the cloud system with the secret by XLConnect. The authorization code is then passed to the user's computer over https to be exchanged for the access token. This mimics the PKCE flow.&#x20;

Because this authorization code can be exchanged only once, it can still be asserted that the token response containing the access and refresh tokens is only ever accessed by the client. Security reviewers are encouraged to use network traffic inspection tools like [Fiddler ](https://www.telerik.com/fiddler)or [Wireshark ](https://www.wireshark.org/)to verify this for themselves.&#x20;

### Secure Storage of Refresh Tokens

One feature of OAuth is the `offline_access` scope that allows the client to login silently with out have to go through the web login every time. This is convenient but requires great care to be taken with the refresh token.&#x20;

These tokens are only stored encrypted using Windows bult-in [System.Security.Cryptography](https://learn.microsoft.com/en-us/dotnet/api/system.security.cryptography?view=netframework-4.8.1) features with User Scope. This means that if the stored tokens were somehow to be stolen they can only be decrypted by the logged in user. This is an extra layer on top of normal Windows BitLocker encryption to ensure the user cannot accidentally share these tokens.&#x20;

## Data Flows Directly from API to User Computer

Data goes directly from a queried API to the client's computer without passing over XLConnect servers.&#x20;

All traffic is over https only, with TLS 1.2 enforced. This is currenlty the highest supported standard.

## Data Lake Storage

For certain applications, the data is too much to be pulled every time, for instance journal transactions. In that case the workbooks can store data locally in the Data Lake.&#x20;

This is only used when required, please check with you developer if the `file` module is being used.&#x20;

Data is stored in 'libraries' that point to a specific folder. Javascript code can only access data at or below these folders. This is a security design to prevent javascript code from being able to scan a user's computer for files. The 'default' library can be accessed at:

&#x20;`C:\Users\<UserName>\AppData\Local\XLConnect\Data`

**Please note**: If used, this is currently not encrypted by XLConnect for transparency reasons and to enable users to use this data for other applications like Python and Power Query etc. The harddrive itself will still be encrypted by BitLocker. If a laptop would fall into the wrong hands the data would still be encrypted.&#x20;

## The Add-In

As the cloud components do not hold any sensitive data, the Add-In must be very secure.&#x20;

The Add-In is written in C# and delivered though a MSI installer that is signed through Extended Validation Code Signing Certificate. This certificate can be inspected by right-clicking the MSI then selecting properties:&#x20;

<figure><img src="../.gitbook/assets/image (47).png" alt=""><figcaption><p>Show properties of installer MSI</p></figcaption></figure>

<figure><img src="../.gitbook/assets/image (48).png" alt=""><figcaption><p>Steps to show EV certificate</p></figcaption></figure>

These certificates are only granted after a rigorous validation process and require a physical token with a password that are both closely guarded. Files without this signature will be flagged by Windows with the [Windows Defender SmartScreen](https://learn.microsoft.com/en-us/windows/security/operating-system-security/virus-and-threat-protection/microsoft-defender-smartscreen/) to warn you about running unsigned files.

All this ensures that the code you run is from XLConnect B.V. In itself this does not prevent XLConnect from shipping particular updates, it does however create a ironclad trail. In the case of suspicion it is technically not difficult to read the .NET code and assert the code is safe. If a party requires a code review or escrow arrangement we can provide that. In the end, a relationship with a software vendor, be it desktop or cloud, requires some trust. This is a solid as it gets.&#x20;

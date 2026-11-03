# Standalone mode

XLConnect Desktop can run a Javascript Studio script **outside Excel, without an interactive login** — for example from the Windows Task Scheduler on a recurring schedule. This is how you automate a daily data pull, a scheduled reconciliation, or any hands-off job.

You can schedule runs on your own computer with your normal personal 'interactive' account, as that will refresh regularly with daily use.

When you want to run the script on another computer where a user does not log in regularly, that type of login can run out, then it is better to setup a standalone login that authenticates as a **service account (NPA)** that is tied to a single database, rather than as a person. It stays valid for as long as that database is active.

{% hint style="info" %}
Standalone runs require the **`desktop.standalone`** feature on the database. It is a feature of the _database_, so a consultant can set it up for a specific client database. Enable it from the XLConnect portal (Database → Features).
{% endhint %}

## How the security works

There are two ways to run a script from the command line:

| Mode            | Command                                    | Runs as                 | Requires                     |
| --------------- | ------------------------------------------ | ----------------------- | ---------------------------- |
| **Interactive** | `XLConnect.Desktop.exe <script>`           | You (browser sign-in)   | `desktop` feature            |
| **Standalone**  | `XLConnect.Desktop.exe <script> <account>` | A named service account | `desktop.standalone` feature |

A standalone credential is stored **encrypted for the Windows account that created it** (Windows DPAPI). This means you must set it up while signed in as the _same_ Windows account the scheduled task will run under — otherwise it cannot be decrypted at run time.

One machine can hold **several** standalone accounts (for example one per client database), each with its own local name.

## Setting up a standalone login

You can set this up from Excel, or on a machine that has no Excel at all.

{% tabs %}
{% tab title="From Excel" %}
1. Turn on **Developer Mode** (ribbon → **Advanced → Developer Mode**).
2. Click **Advanced → Setup Standalone**.
3. Choose the **database** (only databases with the `desktop.standalone` feature are listed — if yours is missing, enable the feature for it in the portal).
4. Enter an **Account** name — this is the local handle you will pass on the command line (e.g. `my-client-db`).
5. Click **Create**.
{% endtab %}

{% tab title="On a machine without Excel" %}
1. Launch **XLConnect Desktop** (run `XLConnect.Desktop.exe` with no arguments) to open the Javascript workbench.
2. Sign in when prompted.
3. Click the **Standalone…** button on the toolbar and follow the same steps as above.
{% endtab %}
{% endtabs %}

{% hint style="warning" %}
Set up the credential **signed in as the Windows account the scheduled task will run under.** The credential is encrypted for that account and cannot be moved to another user.
{% endhint %}

## Running a script

```
XLConnect.Desktop.exe                          Open the Javascript workbench (GUI)
XLConnect.Desktop.exe <script>                 Run <script> as you (interactive sign-in)
XLConnect.Desktop.exe <script> <account>       Run <script> as a standalone service account
XLConnect.Desktop.exe --help                   Show usage and list configured accounts
```

Example:

```
XLConnect.Desktop.exe "C:\Jobs\Daily FX Fixer.js" my-client-account
```

## Scheduling it (Windows Task Scheduler)

1. Create a task and set **"Run whether user is logged on or not"**, using the **same Windows account** you set the credential up under.
2. **Action → Start a program:**
   * Program: `C:\Program Files\XLConnect\XLConnect.Desktop.exe`
   * Arguments: `"C:\Jobs\Daily FX Fixer.js" my-client-account --notify failure`
3. Set your schedule (e.g. daily).

## Notifying users of the outcome

Add `--notify` to email the outcome of the run:

```
XLConnect.Desktop.exe <script> <account> --notify failure|success|both|none
```

The email goes to the users with the **Notify** role on the database (tick **Notify** for a user in the portal, on the database's users). When nobody on the database has that role yet, it goes to the **owners** instead.

| Value                 | Sends an email when… |
| --------------------- | -------------------- |
| `failure` _(default)_ | the run fails       |
| `success`             | the run succeeds    |
| `both`                | every run           |
| `none`                | never               |

The email contains a concise summary (database, script, machine, time, and the error on a failure). The full log stays on the machine that ran the job.

## Logs

Each run writes a log file next to the script: `<script>.<yyyyMMdd>.txt`. It records the credential used, the identity it ran as, and — on a failure — the full error. A failed run also adds the error to `<script>.<yyyyMMdd>ERROR.txt`, so failures stand out in the folder listing.

## Testing a script before scheduling

In the Desktop workbench you can run a script **as** a standalone account to validate exactly what the scheduled job will see (its database scope and features): use the **Run as** selector in the toolbar and pick the account, then run.

## Managing and revoking accounts

Open **Advanced → Setup Standalone** (Excel) or the **Standalone…** button (workbench). When accounts already exist you get a management list where you can **add** another or **Remove** one. Removing an account revokes it on the server and deletes the local credential — any scheduled task using it will stop working.

## Troubleshooting

| Message                                                           | Cause / fix                                                                                                                     |
| ----------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| _"That database does not have the 'desktop.standalone' feature."_ | Enable the `desktop.standalone` feature for that database in the portal.                                                        |
| _Standalone login failed / no token_                              | The credential was revoked, or the database it is tied to is no longer active.                                                  |
| _Credential cannot be decrypted_                                  | The credential was created under a different Windows account than the one running the task. Re-create it as the task's account. |
| _No standalone login named '…'_                                   | Run `XLConnect.Desktop.exe --help` to list configured accounts, or set one up.                                                  |
| _"Connection '…' has no valid saved login, and this unattended run can't show a sign-in window."_ | The script uses a cloud system (Xero, Visma, Moss, …) that has no saved login for this Windows account. Sign in to it once interactively as the Windows account the task runs under, for example by running the script from Excel or the Desktop workbench on that machine. |

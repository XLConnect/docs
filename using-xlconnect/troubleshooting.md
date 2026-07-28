---
description: Answers to typical issues
---

# Troubleshooting

## The XLConnect ribbon is missing

Excel sometimes disables add-ins, for example after a crash. To re-enable:

1. In Excel go to `File > Options > Add-ins`.
2. At the bottom, set the `Manage` dropdown to `COM Add-ins` and press `Go...`.
3. Tick `XLConnect` and press `OK`.
4. If XLConnect is not in the list, check `Manage > Disabled Items` and re-enable it there, then repeat step 3. If it's in neither list, [reinstall the add-in](../installing-the-add-in.md).

## "This functionality requires you to be logged in"

You are not logged in, or your subscription is not active. Press `Login` on the ribbon; if the message persists, check your subscription with your owner or via `Settings > Manage Subscription...`.

## Buttons are missing from the ribbon

Buttons you don't have the required role for are hidden. If you expect developer buttons (Javascript Studio, Save to Cloud, ...), enable `Settings > Advanced > Developer Mode` — and check with your subscription owner that your account has the developer role.

## "Connection failed for ..."

The script couldn't authenticate to the cloud system (for example `xero`). Go to `Settings > Advanced > Manage Tokens`, right-click the connection, choose `Delete`, and run again — you'll be taken through the platform's login to make a fresh connection.

## A workbook opens but shows no data

Connected workbooks are usually published empty — press `Run` to pull the data in.

## The script pauses with a countdown

The cloud platform's API rate limit was hit. XLConnect waits out the timeout the platform requests and retries automatically — no action needed, just let it finish.

## Data looks stale after switching company or period

Press `Clear Data`, make your selection, and `Run` again.

## Still stuck?

Use the `Support` group on the ribbon: `Create ticket...` sends your issue to us, `My tickets...` shows the follow-up.

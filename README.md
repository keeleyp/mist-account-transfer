# Mist Account Transfer

Invites a new email address (e.g. your HPE account) as **Super Observer** to every
Juniper Mist org where your existing accounts are **Super User**, across any or all
Mist clouds.

For every API key in `mist_account_transfer.ini`:

1. Finds every org the key can see (`GET /self`; MSP access is expanded with
   `GET /msps/{msp_id}/orgs`).
2. For orgs where the key is Super User (role `admin` at org or MSP level), reads
   `GET /orgs/{org_id}/admins`:
   - email already an admin → **Already admin**, with its current level
   - email has a pending invite → **Already invited**, with its level and expiry
   - otherwise → invite sent with `POST /orgs/{org_id}/invites`, role `read` +
     view `super_observer` (Super Observer)
   - also checks whether anyone **other than the key owners** is an accepted Super
     User (org or MSP level) on the org. If not, the **Delete Risk** column says
     `ONLY SUPER USER` (red in Excel, and listed as a warning on screen): deleting
     your old account(s) from that org would delete the org, so add another Super
     User first. All key owners (from `/self`) are treated as going away, so having
     both your Juniper and Mist accounts as Super User does not count as safe.
     Pending invites are not counted, as they can expire unused.
   - if such an org is in an **MSP**, the MSP's own Super Users may keep it alive. The
     *MSP* column names the MSP and the Delete Risk says what was found:
     - `ONLY LOCAL SUPER USER - ... MSP has N other Super User(s)` (amber): a key is
       Super User on the MSP, so `GET /msps/{msp_id}/admins` was read; the other
       MSP-level Super Users are listed in *MSP Super Users*. Check with them
       before deleting your account(s)
     - `ONLY SUPER USER - ... MSP has no other Super Users` (red)
     - `ONLY LOCAL SUPER USER - ... check the MSP's admins` (amber): no key can read
       the MSP's admins, so check them by hand
     Only MSP-level Super Users count; the MSP admin list also holds org admins
     of its orgs, which don't.
   - for those sole-Super-User orgs it also lists the org's **inventory**
     (`GET /orgs/{org_id}/inventory?vc=true` - APs, switches including each Virtual
     Chassis member, gateways, Mist Edges). Mist won't delete an org until its
     inventory is released, so each device is listed on the *Inventory to Release*
     sheet and the count shows in the *Inventory* column and the on-screen warning.
   - and its **licences** (`GET /orgs/{org_id}/licenses`): every subscription and
     evaluation with type, quantity, dates and active / expired / not started, plus
     licence moves (amendments) into or out of the org. Moved-in licences belong to
     another org and should be returned (unamended) before the org goes. Listed on the
     *Licences to Remove* sheet, summarised in the *Licences* column and the warning.
3. Orgs where no key is Super User are listed with the access level the key has.
   This includes orgs where `/self` says Super User but Mist refuses the org's
   admin list (HTTP 403). That account is really an Org Admin, so the org is listed
   as Not Super User with a note, and no invite is sent.

The same org seen by two keys (e.g. your Juniper and Mist accounts) is handled
once; a key with Super User wins.

## Output

Screen tables, plus an Excel workbook (created, or updated if it exists) - on a
dry run too, where orgs that would get an invite show *Would invite (dry run)*:

| Sheet | Contents |
|---|---|
| Invitations | Cloud, Org ID, Org Name, Email, Status (Invited / Already admin / Already invited), Level, Delete Risk, Other Super Users, MSP, MSP Super Users, Invited On, Last Checked, Checked With |
| Not Super User | Cloud, Org ID, Org Name, Access Level, Key Owner, Last Checked |
| Inventory to Release | One row per device in a sole-Super-User org: Type, Model, Serial, MAC, Name, Site, Connected, VC MAC. Replaced for each org on every run, so released devices drop off |
| Licences to Remove | One row per subscription / evaluation / licence move in a sole-Super-User org: Kind, Status, Subscription ID, Type, Order ID, Quantity, Remaining, Start, End, Amendment ID. Replaced per org each run |

Rows are updated in place on re-runs (keyed by cloud + org ID), so the workbook
always shows the latest state; *Invited On* is kept from the run that sent the invite.

## Setup

```bash
pip install -r requirements.txt
cp mist_account_transfer.ini.example mist_account_transfer.ini
```

Fill in one token per cloud under `[clouds]` (comma-separate several tokens on
one cloud), and `invite_email`, `first_name`, `last_name` under `[settings]`.

## Run

```bash
python3 mist_account_transfer.py --dry-run   # look only: writes the workbook, sends nothing
python3 mist_account_transfer.py             # shows the list, asks before sending
python3 mist_account_transfer.py --yes       # send without asking
python3 mist_account_transfer.py --ini ~/customer_x.ini --dry-run   # another settings file
```

`--ini PATH` uses a different settings file instead of `mist_account_transfer.ini`
next to the script. A relative `output_file` is then saved next to that ini, so
each settings file keeps its own workbook.

Safety: every key and every Super User org is checked before anything is sent. If
any fails, the script stops and lists the failures. If an invite fails partway, it
stops; invites already sent are recorded in the workbook.

## Zscaler / proxies

If you're behind a TLS-inspecting proxy, set `[network] ca_bundle` to your proxy's
root CA PEM file. `http_proxy` / `https_proxy` override the environment variables.
Leave the section out for normal behaviour.

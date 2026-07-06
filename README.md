# F5 BIG-IP iHealth QKview Parser 

Pulls BIG-IP sizing intelligence from one or more iHealth QKViews and writes a
clean multi-sheet Excel report — identity, licensing, provisioning, configuration
totals, and Highest/Average performance figures for CPU, throughput, SSL TPS,
active connections, memory, and DNS across 3 hour / 1 day / 7 day / 30 day
windows.



## What it does

For each QKView ID in `QKVIEWS`:

| Section | Data |
|---|---|
| **Inventory** | hostname, platform, version, build, edition, chassis S/N, license platform ID |
| **Provisioning & License** | provisioned modules + levels, active licensed modules, license headline |
| **Config Totals** | object counts (LTM virtuals, pools, nodes, monitors, profiles, iRules, policies, SNAT, GTM wide-IPs, ASM/APM policies, net objects), with a per-column SUM |
| **Performance** | per-blade per-core CPU usage %, total throughput (bits/s), SSL TPS, active connections, memory used (bytes and percent), DNS request rate — each with Highest + Average across 3 h / 1 d / 7 d / 30 d |
| **Notes** | data-source attribution and methodology |

## Requirements

- Python 3.9+
- `rrdtool` (system package — `apt-get install rrdtool` on Debian/Ubuntu/WSL)
- Python packages: `requests`, `openpyxl` (see `requirements.txt`)
- iHealth API Client ID + Client Secret (generated in the iHealth GUI → Settings)

## Setup

```bash
git clone https://github.com/<your-org>/f5-ihealth-sizing.git
cd f5-ihealth-sizing
sudo apt-get install -y rrdtool
pip install -r requirements.txt
```

Generate iHealth API credentials at <https://ihealth.f5.com> → **Settings**, then export:

```bash
export IHEALTH_CLIENT_ID="..."
export IHEALTH_CLIENT_SECRET="..."
```

The credentials are scoped to your iHealth account — the API only returns QKViews
owned by the account that generated the keys. If a QKView is shared with you
through a support case (visible in the GUI but not your collection), download it
from the GUI and re-upload it under your account first:

```bash
python3 ihealth_sizing.py --upload /path/to/device.qkview
# prints the new QKView ID — add it to QKVIEWS below
```

## Configuration

Edit two things at the top of `ihealth_sizing.py`:

**1. `QKVIEWS` — the QKViews to analyze.** Key = numeric QKView ID (from the
GUI URL `https://ihealth.f5.com/qkview-analyzer/qv/<id>`). Value = friendly label
that appears in the report.

```python
QKVIEWS = OrderedDict([
    ("2xxxxx", "device-01"),
    ("2xxxxx", "device-02"),
])
```

**2. `OUTPUT_XLSX`** — output filename (defaults to `ihealth_sizing_report.xlsx`).

Everything else (`PERF_METRICS`, `CONFIG_COUNTS`, `INTERVALS`) ships with sensible
defaults tuned to TMOS 17.x but can be overridden if you hit a device whose DS
naming differs (see *Troubleshooting* below).

### Optional environment overrides

| Variable | Default | Purpose |
|---|---|---|
| `IHEALTH_CLIENT_ID` | *(required)* | OAuth2 client ID |
| `IHEALTH_CLIENT_SECRET` | *(required)* | OAuth2 client secret |
| `IHEALTH_API_BASE` | `https://ihealth-api.f5.com/qkview-analyzer/api` | API host |
| `IHEALTH_TOKEN_URL` | `https://identity.account.f5.com/oauth2/ausp95ykc80HOU7SQ357/v1/token` | OAuth token endpoint |
| `IHEALTH_UA` | `ihealth-sizing/<ver>` | User-Agent string |

## Usage

```bash
# Full run — analyze every QKView in QKVIEWS, write Excel
python3 ihealth_sizing.py

# Discovery — list /var/tmp/qkview-rrd RRD files + their DS names per QKView
# (use this when adapting PERF_METRICS to a TMOS version you haven't seen)
python3 ihealth_sizing.py --discover

# Upload a local .qkview file to your iHealth account
python3 ihealth_sizing.py --upload /path/to/device.qkview
```

Progress prints to stderr (one section per QKView). Expect ~30 – 90 seconds per
QKView depending on size; a run of 14 mixed appliances + chassis usually takes
5 – 15 minutes end-to-end.

## How it works

All data comes from the **current iHealth REST API** at `ihealth-api.f5.com`.
The 2023 clouddocs reference (`ihealth2-api.f5.com`) is deprecated.

| Data | Endpoint | Method |
|---|---|---|
| Identity | `GET /qkviews/<id>/diagnostics.json?set=hit` | parse `system_information` + `version` |
| Active modules | `/config/bigip.license` (Files API, octet-stream) | parse `active module :` lines |
| Provisioning | `/config/bigip_base.conf` (Files API) | regex `sys provision <name> { level <lvl> }` |
| Config totals | `/config/bigip.conf` + `/config/bigip_base.conf` (Files API) | count column-0 tmsh stanzas |
| Performance | `/var/tmp/qkview-rrd/*.xml.gz` (Files API) | gunzip → `rrdtool restore` → `rrdtool graph` PRINT VDEF |

### Performance methodology

iHealth's performance graphs are powered by per-blade RRD data exported to
`rrdtool dump` XML inside the qkview. The script:

1. Downloads the relevant `bladeNcpu.xml.gz`, `throughput.xml.gz`,
   `connections.xml.gz`, `memory.xml.gz` files.
2. `gunzip` + `rrdtool restore` → binary RRD.
3. For each metric + window, runs `rrdtool graph` with a DEF/CDEF/VDEF chain
   that produces "Highest" and "Average" for that window.

**CPU usage %** is computed as `(sum(all modes) − idle) / sum(all modes) × 100`
inside `rrdtool` via CDEF, where modes = `idle + user + system + iowait + irq +
softirq + niced + stolen`. This is self-normalizing across kernel HZ values
(100/250/1000) and the per-step ratio is taken **before** VDEF aggregation, so
windowed Highest/Average are mathematically correct.

**Memory used %** is computed as `(Rtmmused + Rotherused) / Rtotal × 100`
inside `rrdtool` CDEF, matching the iHealth GUI's percent display.

**Throughput / SSL TPS / Active Connections / DNS** use straight `MAX` and
`AVERAGE` consolidation from the relevant DSes (rrdtool returns per-second rates
automatically for COUNTER-type DSes).

Source attribution for every performance row is recorded in the workbook (which
RRD file and which DSes were used) so the math is auditable.

### Multi-blade chassis support

The CPU metric uses `files_regex: ^blade\d+cpu\.xml\.gz$` and iterates **every**
populated blade slot. A 4-blade chassis with 12 cores per blade produces
48 CPU rows (Blade 1 / Core 0 … Blade 4 / Core 11). Memory uses the rollup
(`R*`) DSes for whole-device totals — no per-blade double-counting.

## Troubleshooting

### `HTTP 404` on every QKView, but the QKView is visible in the iHealth GUI
Two causes, in order:

1. **Wrong API host.** The current API is `ihealth-api.f5.com` (not the legacy
   `ihealth2-api.f5.com` from the 2023 clouddocs). Make sure your
   `IHEALTH_API_BASE` env var, if set, points at the current host.
2. **Ownership scope.** The API only returns QKViews owned by the account that
   issued your Client ID/Secret. A QKView shared with you via support case is
   visible in the GUI but not in your API collection. Re-upload it under your
   account with `--upload`.

Quick probe to confirm which case you're in:

```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  -H "Accept: application/vnd.f5.ihealth.api+json" \
  "https://ihealth-api.f5.com/qkview-analyzer/api/qkviews"
```

If this returns `{"id":[]}` your collection is empty (ownership issue). If it
returns IDs but a specific one 404s, that QKView genuinely doesn't exist in
your account.

### `HTTP 503` or connection drops mid-run
Transient iHealth gateway issues. The script retries 500/502/503/504 and
network drops with exponential backoff (3s → 6s → 12s … capped at 60s).
A run interrupted at auth can be re-run from scratch with no side effects.

### A metric shows `N/A — no DS matched [...]`
The device's TMOS version uses different DS names than the defaults. Run:

```bash
python3 ihealth_sizing.py --discover
```

That dumps every `/var/tmp/qkview-rrd/*.xml.gz` file and its DS names per
QKView. Update the `ds_match` substrings (or `files_regex`) in `PERF_METRICS`
at the top of the script.

### A metric shows `N/A — no qkview-rrd source found`
The expected RRD isn't in this QKView. Common cause: the module isn't
provisioned on the device (e.g., DNS Requests N/A on a device without
GTM/DNS provisioned). This is the correct, honest result — not a bug.

### `ConnectionTimeout` to `identity.account.f5.com`
Network reachability problem (VPN flap, WSL2 NIC hiccup, corporate proxy).
The host is reachable from any normal internet egress. Confirm with
`curl -v https://identity.account.f5.com/`.

### Excel formula errors / weird sheet output
Excel reopens with `=SUM(...)` on the Config Totals sheet; if you see `#REF!`
or empty totals, your spreadsheet app may not have recalculated. Re-open or
press F9.

## Limitations

- **Read-only.** This tool does not modify any device configuration.
- **Performance metrics depend on the device's TMOS RRD schema.** Defaults are
  tuned to TMOS 17.x; older versions may need `ds_match` adjustments
  (`--discover` shows what to change).
- **No platform-spec percent calculations** for throughput / SSL TPS / connections.
  Raw absolute values are reported; percent-of-platform-max requires a hardware
  spec table not currently built in.
- **Single-account scoping.** API credentials see only their own QKView collection.
- **PII.** QKViews contain device hostnames, IPs, certificate metadata, and
  configuration. Treat the output Excel as the same sensitivity as the QKView
  itself.

## Project layout

```
.
├── README.md
├── LICENSE
├── CHANGELOG.md
├── requirements.txt
├── .gitignore
└── ihealth_sizing.py     # the script
```

## License

MIT — see `LICENSE`.

## Acknowledgments

Built against the public iHealth REST API documented at
<https://clouddocs.f5.com/api/ihealth/> (legacy) and the live Swagger at
<https://ihealth-api.f5.com/qkview-analyzer/api/docs/index.html>.

Not affiliated with or endorsed by F5, Inc. The iHealth API itself is not
covered by F5 support; use at your own discretion.

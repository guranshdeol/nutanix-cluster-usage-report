<div align="center">

# <img src="https://cdn.simpleicons.org/nutanix/7855FA" alt="" height="34" align="top"> Nutanix Cluster Utilization Report

### 30 days of CPU, memory and storage for every cluster, in one Excel workbook

[![Python](https://img.shields.io/badge/python-3.9%2B-blue?logo=python&logoColor=white)](https://www.python.org/downloads/)
[![Nutanix Prism v4](https://img.shields.io/badge/Nutanix-Prism%20v4-red)](https://developers.nutanix.com/)
[![Read-only](https://img.shields.io/badge/cluster%20access-read--only-brightgreen)](#apis-used)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)](#setup)

*Log in to Prism Central → pick clusters → get a summary sheet plus one charted sheet per cluster.*

</div>

---

An interactive Python script that connects to **Nutanix Prism Central** using the **v4.0 REST APIs** and generates a formatted Excel report of **CPU, memory and storage usage over the last 30 days** for every cluster it manages.

## Features

- Interactive prompts, with no config files or hard-coded credentials
- Local user authentication (e.g. `admin`) with a hidden password prompt
- Optional SSL certificate verification (skip it for self-signed certificates)
- Report on all registered clusters or pick specific ones
- Excel output with:
  - **Summary** sheet: average and peak CPU %, average and peak memory %, storage used, capacity, used % and 30-day growth per cluster
  - **One sheet per cluster**: daily averages and peaks, plus line charts for CPU, memory and storage
  - Styled headers, zebra striping, frozen panes and green/yellow/red colour coding for utilization

## How it works

```mermaid
flowchart LR
  login["Prism Central login<br/>(hidden password prompt)"] --> list["List clusters<br/>clustermgmt v4.0"]
  list --> pick["Pick clusters or all<br/>(PC entry skipped)"]
  pick --> stats["Daily stats for 30 days<br/>AVG and MAX"]
  stats --> book["Excel workbook<br/>Summary + one sheet per cluster"]
```

Every call is a `GET`. The script never changes anything on Prism Central or the clusters.

## Requirements

- Python 3.9+
- Network access to Prism Central on port 9440 (default)
- A Prism Central local user that can read cluster configuration and stats

## Setup

```bash
git clone https://github.com/guranshdeol/nutanix-cluster-usage-report.git
cd nutanix-cluster-usage-report
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
python cluster_usage_report.py
```

Example session:

```
Prism Central IP/FQDN: pc.example.local
Port [9440]:
Local admin username [admin]:
Password:
Verify SSL certificate? (y/N):

Found 2 cluster(s):
  1. Cluster-A
  2. Cluster-B

Clusters to include (e.g. 1,3) or 'all' [all]:
Fetching stats for Cluster-A ... 30 day(s)
Fetching stats for Cluster-B ... 30 day(s)

Output file [cluster_usage_YYYYMMDD_HHMM.xlsx]:
Report saved to /path/to/cluster_usage_YYYYMMDD_HHMM.xlsx
```

## APIs used

| Purpose | Endpoint |
|---|---|
| List clusters | `GET /api/clustermgmt/v4.0/config/clusters` |
| Cluster stats | `GET /api/clustermgmt/v4.0/stats/clusters/{extId}` |

Stats are requested with a daily sampling interval (`$samplingInterval=86400`), once with `$statType=AVG` and once with `$statType=MAX`. The metrics used are:

- `hypervisorCpuUsagePpm`: CPU usage (parts per million, converted to %)
- `aggregateHypervisorMemoryUsagePpm`: memory usage (parts per million, converted to %)
- `storageUsageBytes` / `storageCapacityBytes`: storage used and capacity (converted to GiB)

The Prism Central entry that appears in the cluster list is skipped automatically.

## Notes

- Stats cover the last 30 days in UTC. Clusters registered more recently show fewer days.
- The password is only ever read at runtime and is never stored.

---

### More Nutanix tools

- [Multi-Cluster Command Runner](https://github.com/guranshdeol/nutanix-multi-cluster-command-runner): run one shell command on every cluster
- [Prism Central Inventory (PowerShell)](https://github.com/guranshdeol/nutanix-cluster-inventory-v4-ps): full estate inventory in one workbook
- [VLAN Migrator](https://github.com/guranshdeol/nutanix-vlan-migrator): Basic to Advanced subnet migration, with validation
- [SSH Consent Banner](https://github.com/guranshdeol/nutanix-ahv-cvm-custom-banner): pre-login banner on CVMs and AHV hosts

Made by [**Guransh Deol**](https://github.com/guranshdeol)

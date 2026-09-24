# Nutanix Cluster Utilization Report

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

## Requirements

- Python 3.9+
- Network access to Prism Central on port 9440 (default)
- A Prism Central local user that can read cluster configuration and stats

## Setup

```bash
git clone <this-repo-url>
cd <repo-folder>
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

Made by **Guransh Deol**

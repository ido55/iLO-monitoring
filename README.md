# iLO Monitoring Script

A Windows desktop application for collecting hardware health information from HPE iLO interfaces. 
Uses the HPE `hpqlocfg` utility to query a list of servers, tracks collection progress, matches IP addresses to site names, and exports the results to a formatted Excel workbook.

## Features

- Collects server identity, firmware, and component health status (fan status, temperature etc.).
- Shows per-server progress and collection errors in a Tkinter interface.
- Associates each iLO address with a site name.
- Exports results to an `.xlsx` file.
- Includes an offline preview using synthetic data.

## Requirements

- Windows
- Python 3.10 or newer with Tkinter
- HPE `hpqlocfg` for live collection

The HPE utility is distributed separately and is not included in this repository. The application can be opened and previewed without it. `pandas` and `XlsxWriter` are installed by the setup steps below.

## Setup and launch

From the project folder, run in Powershell:

----------------------------------------------------------
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python GUI_iLO_Script.py
----------------------------------------------------------

The application opens a desktop window. Select your target list and site map, enter the iLO account credentials, and click **Start Process**. When collection finishes, click **Export to Excel**. Reports and collector logs are saved in the `reports` folder.

To try the interface and report flow without connecting to any server, click **Preview sample (offline)**, then **Export to Excel**. The sample uses synthetic data and makes no network requests.

## Configure targets and sites

### Target list

Edit `ScriptFiles/iloListTestNew.txt`. Add one authorized iLO IP address per line. Blank lines and lines beginning with `#` are ignored. The supplied file has no targets, so add addresses before starting a live collection. You can also choose a different target file in the application.

```text
10.0.0.10
10.0.0.11
```

### Site mapping

Edit `ScriptFiles/ILOs_And_Sites.csv` using UTF-8 CSV format and the column headings `IP` and `Site`:

```csv
IP,Site
10.0.0.10,Datacenter A
10.0.0.11,Datacenter B
```

An address without a matching row is still collected; its site field will be blank. You can choose a different mapping file in the application.

### HPE utility path

The application checks common installation folders for `hpqlocfg.exe`. If it is installed elsewhere, set `HPQLOCFG_PATH` in PowerShell before launching:

```powershell
$env:HPQLOCFG_PATH = "C:\path\to\hpqlocfg.exe"
.\.venv\Scripts\python GUI_iLO_Script.py
```

Use a utility version compatible with the iLO hardware and firmware you manage. `ScriptFiles/Get-ALL.xml` contains the RIBCL query sent by `hpqlocfg`.

## Security

Use an account with only the permissions needed to read health information, and run the application on a trusted workstation. The legacy `hpqlocfg` interface receives credentials as command-line arguments, so a sufficiently privileged local user may be able to inspect them while collection runs. The generated collector log can contain server response data; protect it and the report as operational data.

Only collect from systems you are authorized to monitor. Review target files, reports, and logs before sharing or publishing them.

## Credits

`ScriptFiles/Get-ALL.xml` is an HPE sample script and retains its HPE copyright and attribution notice. HPE software is not bundled with this project.

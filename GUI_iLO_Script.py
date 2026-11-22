"""iLO Monitoring Script — refreshed from the original 2022 application.

The application keeps the original flow: enter credentials, collect each
configured iLO through hpqlocfg, track progress, and export an Excel report.
"""
from __future__ import annotations

import csv
import os
import queue
import subprocess
import threading
import tkinter as tk
import xml.etree.ElementTree as ET
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
SCRIPT_FILES = BASE_DIR / "ScriptFiles"
TARGET_FILE = SCRIPT_FILES / "iloListTestNew.txt"
XML_SCRIPT = SCRIPT_FILES / "Get-ALL.xml"
SAMPLE_RESPONSE = SCRIPT_FILES / "sample-response.xml"
SITE_FILE = SCRIPT_FILES / "ILOs_And_Sites.csv"
OUTPUT_DIR = BASE_DIR / "reports"
OUTPUT_FILE = OUTPUT_DIR / "Output.xlsx"
LOG_FILE = OUTPUT_DIR / "hpqlocfg.log"

HEADERS = [
    "IP address", "Site Name", "Server Name", "iLO Product", "iLO Version",
    "BIOS Hardware Status", "Fans Status", "Fans Redundancy",
    "Temperature Status", "Power Supplies Status", "Power Supplies Redundancy",
    "Processor Status", "Memory Status", "Storage Status", "Collection Error",
]

# Kept as module state to make the original callbacks straightforward to follow.
counter = 0
NUMiLOs = 0
iLOCoList: list[str] = []
results: list[dict[str, str]] = []
events: queue.Queue = queue.Queue()


def read_targets() -> list[str]:
    """Read target addresses; blank lines and comments are harmless."""
    targets = []
    for number, line in enumerate(TARGET_FILE.read_text(encoding="utf-8-sig").splitlines(), 1):
        address = line.split("#", 1)[0].strip()
        if not address:
            continue
        if not all(char.isdigit() or char in ".:" or char.lower() in "abcdef" for char in address):
            raise ValueError(f"Invalid address on line {number}: {address!r}")
        if address not in targets:
            targets.append(address)
    if not targets:
        raise ValueError(f"No iLO addresses are listed in {TARGET_FILE}")
    return targets


def get_site_map() -> dict[str, str]:
    """Load the original IP-to-site mapping from a simple CSV file."""
    sites: dict[str, str] = {}
    with SITE_FILE.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or not {"IP", "Site"}.issubset(reader.fieldnames):
            raise ValueError("ILOs_And_Sites.csv must have IP and Site columns")
        for row in reader:
            ip = (row.get("IP") or "").strip()
            if ip:
                sites[ip] = (row.get("Site") or "").strip()
    return sites


def find_hpqlocfg() -> Path | None:
    """Find the HPE utility, or allow an explicit path through an environment variable."""
    configured = os.environ.get("HPQLOCFG_PATH", "").strip().strip('"')
    candidates = [Path(configured)] if configured else []
    candidates += [
        Path(r"C:\Program Files\Hewlett Packard Enterprise\HPQLOCFG\hpqlocfg.exe"),
        Path(r"C:\Program Files (x86)\Hewlett Packard Enterprise\HP Lights-Out Configuration Utility\hpqlocfg.exe"),
        Path(r"C:\Program Files (x86)\Hewlett Packard Enterprise\HP Lights-Out Configuration Utility\hpqlocfg"),
    ]
    return next((path for path in candidates if path.is_file()), None)


def _tag(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1].upper()


def parse_health_response(xml_text: str, address: str, site: str) -> dict[str, str]:
    """Convert an hpqlocfg RIBCL response into one of the original report rows."""
    row = dict.fromkeys(HEADERS, "")
    row["IP address"], row["Site Name"] = address, site
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        row["Collection Error"] = f"Invalid XML: {exc}"
        return row

    for element in root.iter():
        name = _tag(element)
        if name == "GET_FW_VERSION":
            row["iLO Version"] = element.get("FIRMWARE_VERSION", "").strip()
        elif name == "SERVER_NAME":
            row["Server Name"] = element.get("VALUE", "").strip()
        elif name == "PRODUCT_NAME":
            row["iLO Product"] = element.get("VALUE", "").strip()
        elif name in {"BIOS_HARDWARE", "FANS", "TEMPERATURE", "POWER_SUPPLIES", "PROCESSOR", "MEMORY", "STORAGE"}:
            label = {
                "BIOS_HARDWARE": "BIOS Hardware Status", "FANS": "Fans Status",
                "TEMPERATURE": "Temperature Status", "POWER_SUPPLIES": "Power Supplies Status",
                "PROCESSOR": "Processor Status", "MEMORY": "Memory Status", "STORAGE": "Storage Status",
            }[name]
            row[label] = element.get("STATUS", "").strip()
            redundancy = element.get("REDUNDANCY", "").strip()
            if redundancy:
                row["Fans Redundancy" if name == "FANS" else "Power Supplies Redundancy"] = redundancy
    return row


def getDataToTextFile(username: str, password: str, hpqlocfg: Path) -> None:
    """Collect the configured iLOs one at a time, retaining the original workflow."""
    global counter, results
    results = []
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    sites = get_site_map()
    log_lines = []
    for index, address in enumerate(iLOCoList, 1):
        command = [str(hpqlocfg), "-f", str(XML_SCRIPT), "-s", address,
                   "-t", f'user="{username}",password="{password}"', "-l", str(LOG_FILE)]
        try:
            completed = subprocess.run(command, cwd=hpqlocfg.parent, capture_output=True,
                                       text=True, timeout=60, check=False)
            response = completed.stdout
            if not response.strip() and completed.stderr.strip():
                response = completed.stderr
            row = parse_health_response(response, address, sites.get(address, ""))
            if completed.returncode:
                row["Collection Error"] = row.get("Collection Error") or f"hpqlocfg exit code {completed.returncode}"
            # Keep raw output in the local report folder for troubleshooting, as before.
            log_lines.append(f"--- {address} (exit {completed.returncode}) ---\n{response}\n")
        except subprocess.TimeoutExpired:
            row = dict.fromkeys(HEADERS, "")
            row.update({"IP address": address, "Site Name": sites.get(address, ""),
                        "Collection Error": "Collection timed out after 60 seconds"})
            log_lines.append(f"--- {address}: timed out ---\n")
        except OSError as exc:
            row = dict.fromkeys(HEADERS, "")
            row.update({"IP address": address, "Site Name": sites.get(address, ""),
                        "Collection Error": f"Could not run hpqlocfg: {exc}"})
            log_lines.append(f"--- {address}: {exc} ---\n")
        results.append(row)
        counter = index
        events.put(("progress", index, len(iLOCoList)))
    # Never write credentials into logs or reports.
    LOG_FILE.write_text("".join(log_lines), encoding="utf-8", errors="replace")
    username = password = ""
    events.put(("finished",))


def Text_To_Excel() -> None:
    """Export collected rows into the familiar color-coded Excel workbook."""
    if not results:
        messagebox.showinfo("No results", "Run the collection before exporting.")
        return
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    frame = pd.DataFrame(results, columns=HEADERS)
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
            frame.to_excel(writer, sheet_name="list", index=False)
            workbook = writer.book
            worksheet = writer.sheets["list"]
            header_format = workbook.add_format({"bold": True, "bg_color": "#DEDE00", "text_wrap": True})
            for column, heading in enumerate(HEADERS):
                worksheet.write(0, column, heading, header_format)
                values = [len(str(row.get(heading, ""))) for row in results]
                worksheet.set_column(column, column, min(max([len(heading), *values]) + 2, 40))
            colors = {
                "ok": "#82DD69", "Failed": "#FF0000", "Link Down": "#FF0000",
                "Not": "#E79603", "Degraded": "#E79603", "Redundant": "#F2EC00",
            }
            for value, color in colors.items():
                worksheet.conditional_format(1, 0, len(results), len(HEADERS) - 1,
                    {"type": "text", "criteria": "containing", "value": value,
                     "format": workbook.add_format({"bg_color": color, "font_color": "#000000"})})
            worksheet.freeze_panes(1, 2)
            worksheet.autofilter(0, 0, len(results), len(HEADERS) - 1)
    except ImportError:
        csv_path = OUTPUT_FILE.with_suffix(".csv")
        frame.to_csv(csv_path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("Report exported", f"Excel support is unavailable. CSV saved to:\n{csv_path}")
        return
    messagebox.showinfo("Report exported", f"Report saved to:\n{OUTPUT_FILE}")
    if os.name == "nt":
        try:
            os.startfile(OUTPUT_FILE)  # type: ignore[attr-defined]
        except OSError:
            pass


def _process_events() -> None:
    """Apply worker-thread messages on Tk's main thread."""
    try:
        while True:
            event = events.get_nowait()
            if event[0] == "progress":
                _, current, total = event
                P1.configure(value=current, maximum=total)
                L3.configure(text=f"Checking iLO {current} of {total}", bg="#FFF2A8")
            elif event[0] == "error":
                messagebox.showerror("Collection error", event[1])
            elif event[0] == "finished":
                StartButton.configure(state="normal")
                ExportButton.configure(state="normal")
                L3.configure(text="Script is Done", bg="#90EE90")
    except queue.Empty:
        pass
    root.after(100, _process_events)


def TDF() -> None:
    """Start collection without blocking the GUI."""
    username, password = Entry1.get().strip(), Entry2.get()
    hpqlocfg = find_hpqlocfg()
    try:
        targets = read_targets()
        get_site_map()
    except (OSError, ValueError) as exc:
        messagebox.showerror("Configuration error", str(exc))
        return
    if not username or not password:
        messagebox.showerror("Missing credentials", "Enter the iLO username and password first.")
        return
    if hpqlocfg is None:
        messagebox.showerror("hpqlocfg not found", "Install the HPE utility or set HPQLOCFG_PATH to its executable.")
        return
    global iLOCoList, NUMiLOs, counter
    iLOCoList, NUMiLOs, counter = targets, len(targets), 0
    Entry2.delete(0, tk.END)
    StartButton.configure(state="disabled")
    ExportButton.configure(state="disabled")
    P1.configure(value=0, maximum=NUMiLOs)
    L3.configure(text="Script in Progress", bg="#FFF2A8")

    def worker() -> None:
        try:
            getDataToTextFile(username, password, hpqlocfg)
        except Exception as exc:
            events.put(("error", str(exc)))
            events.put(("finished",))

    threading.Thread(target=worker, name="getDataToTextFile", daemon=True).start()


def preview_sample() -> None:
    """Load synthetic response data to preview the app and export without networking."""
    global results
    try:
        sample_xml = SAMPLE_RESPONSE.read_text(encoding="utf-8-sig")
    except OSError as exc:
        messagebox.showerror("Sample unavailable", str(exc))
        return
    results = [parse_health_response(sample_xml, "sample-target", "Sample site")]
    P1.configure(value=1, maximum=1)
    L3.configure(text="Sample Preview Ready — no connection made", bg="#C0F8D8")
    ExportButton.configure(state="normal")


def _browse_targets() -> None:
    filename = filedialog.askopenfilename(initialdir=SCRIPT_FILES, title="Select iLO target list")
    if filename:
        global TARGET_FILE
        TARGET_FILE = Path(filename)


def _browse_sites() -> None:
    filename = filedialog.askopenfilename(initialdir=SCRIPT_FILES, title="Select IP-to-site CSV")
    if filename:
        global SITE_FILE
        SITE_FILE = Path(filename)


def launch_gui() -> None:
    """Build and run the original Tkinter workflow."""
    global root, Entry1, Entry2, StartButton, ExportButton, PreviewButton, P1, L3
    # Main window: the original login -> start -> progress -> export flow is retained.
    root = tk.Tk()
    root.title("iLO Health Monitoring")
    root.geometry("900x570")
    root.minsize(720, 480)
    root.resizable(True, True)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    main = ttk.Frame(root, padding=24)
    main.grid(sticky="nsew")
    main.columnconfigure(1, weight=1)

    ttk.Label(main, text="iLO Health Monitoring", font=("Segoe UI", 24, "bold")).grid(
    row=0, column=0, columnspan=3, sticky="w", pady=(0, 20))
    ttk.Label(main, text="Enter the shared iLO account used for this authorized target list.",
          wraplength=650).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 16))
    ttk.Label(main, text="Username").grid(row=2, column=0, sticky="w", pady=6)
    Entry1 = ttk.Entry(main)
    Entry1.grid(row=2, column=1, columnspan=2, sticky="ew", padx=8, pady=6)
    ttk.Label(main, text="Password").grid(row=3, column=0, sticky="w", pady=6)
    Entry2 = ttk.Entry(main, show="*")
    Entry2.grid(row=3, column=1, columnspan=2, sticky="ew", padx=8, pady=6)
    ttk.Label(main, text="Security note: hpqlocfg receives credentials as command-line arguments; a privileged local user may be able to inspect them.",
              wraplength=720, foreground="#8A3B00").grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 12))
    ttk.Label(main, text=f"Targets: {TARGET_FILE}", wraplength=650).grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 4))
    ttk.Button(main, text="Choose…", command=_browse_targets).grid(row=5, column=2, sticky="e", pady=(8, 4))
    ttk.Label(main, text=f"Site map: {SITE_FILE}", wraplength=650).grid(row=6, column=0, columnspan=2, sticky="w", pady=4)
    ttk.Button(main, text="Choose…", command=_browse_sites).grid(row=6, column=2, sticky="e", pady=4)
    L3 = tk.Label(main, text="Start Script", font=("Segoe UI", 16, "bold"), bg="#E8E8E8", padx=10, pady=8)
    L3.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(22, 12))
    P1 = ttk.Progressbar(main, orient="horizontal", mode="determinate")
    P1.grid(row=8, column=0, columnspan=3, sticky="ew", pady=8)
    buttons = ttk.Frame(main)
    buttons.grid(row=9, column=0, columnspan=3, sticky="e", pady=14)
    StartButton = ttk.Button(buttons, text="Start Process", command=TDF)
    StartButton.pack(side="left", padx=5)
    PreviewButton = ttk.Button(buttons, text="Preview sample (offline)", command=preview_sample)
    PreviewButton.pack(side="left", padx=5)
    ExportButton = ttk.Button(buttons, text="Export to Excel", command=Text_To_Excel, state="disabled")
    ExportButton.pack(side="left", padx=5)
    ttk.Label(main, text=f"Reports and collector logs are saved in: {OUTPUT_DIR}", wraplength=650).grid(
        row=10, column=0, columnspan=3, sticky="w", pady=6)

    root.after(100, _process_events)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()

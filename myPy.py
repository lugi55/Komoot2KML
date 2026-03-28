import os
import subprocess
import sys
import simplekml
import glob
import gpxpy
import re
import xml.etree.ElementTree as ET
import shutil

EMAIL    = ""
PASSWORD = ""
BASE_DIR = "gpx_files"

SPORT_MAP = {
    "racebike":       "cycling",
    "touringbicycle": "cycling",
    "hike":           "hiking",
    "mountaineering": "hiking",
    "skitour":        "winter",
    "skialpin":       "winter",
    "other":          "other",
}

CATEGORY_MIN_KM = {
    "cycling": 30.0,
}

CATEGORY_COLORS = {
    "cycling": simplekml.Color.purple,
    "hiking":  simplekml.Color.red,
    "winter":  simplekml.Color.orange,
    "other":   simplekml.Color.white,
    "unknown": simplekml.Color.yellow,
}

CLR_GREEN  = "\033[92m"
CLR_YELLOW = "\033[93m"
CLR_RED    = "\033[91m"
CLR_CYAN   = "\033[96m"
CLR_RESET  = "\033[0m"

def check_unknown_sports():
    """Lists all tours and warns about any sport types not in SPORT_MAP."""
    print(f"{CLR_CYAN}Checking for unknown sport types...{CLR_RESET}\n")
    cmd = [
        sys.executable, "-m", "komootgpx",
        f"--mail={EMAIL}",
        f"--pass={PASSWORD}",
        "--list-tours",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"{CLR_RED}ERROR: {result.stderr.strip()}{CLR_RESET}")
        return
    unknown = {}  # sport -> list of tour names
    known_counts = {}  # sport -> count
    for line in result.stdout.splitlines():
        match = re.search(r'=>\s+(.+?)\s+\((\w+);', line)
        if not match:
            continue
        name  = match.group(1)
        sport = match.group(2).lower()
        if sport not in SPORT_MAP:
            unknown.setdefault(sport, []).append(name)
        else:
            known_counts[sport] = known_counts.get(sport, 0) + 1
    print(f"{'SPORT':<20} {'CATEGORY':<12} {'COUNT'}")
    print("-" * 45)
    for sport, count in sorted(known_counts.items()):
        category = SPORT_MAP[sport]
        print(f"{sport:<20} {category:<12} {count}")
    if unknown:
        print(f"\n{CLR_YELLOW}WARNING: Unknown sport types found:{CLR_RESET}")
        print("-" * 45)
        for sport, names in unknown.items():
            print(f"{CLR_YELLOW}  '{sport}' ({len(names)} tours):{CLR_RESET}")
            for name in names:
                print(f"    - {name}")
        print(f"\n{CLR_YELLOW}Add these to SPORT_MAP to categorize them correctly.{CLR_RESET}")
    else:
        print(f"\n{CLR_GREEN}All sport types are covered in SPORT_MAP.{CLR_RESET}")


def batch_convert_all_categories(base_gpx_dir, base_kml_dir):
    """Converts GPX files from category subfolders to KML, preserving folder structure."""
    print(f"\n{'FILE':<35} | {'CATEGORY':<12} | {'DISTANCE':<10}")
    print("-" * 65)
    for category, color in CATEGORY_COLORS.items():
        gpx_dir = os.path.join(base_gpx_dir, category)
        kml_dir = os.path.join(base_kml_dir, category)
        if not os.path.exists(gpx_dir):
            continue
        os.makedirs(kml_dir, exist_ok=True)
        gpx_files = glob.glob(os.path.join(gpx_dir, "*.gpx"))
        for gpx_path in gpx_files:
            base_name = os.path.basename(gpx_path)
            stem      = os.path.splitext(base_name)[0]
            id_match  = re.search(r'(\d+)', stem)
            id_str    = id_match.group(1) if id_match else stem
            kml_name  = id_str + ".kml"
            kml_path  = os.path.join(kml_dir, kml_name)
            if os.path.exists(kml_path):
                continue
            try:
                with open(gpx_path, 'r', encoding='utf-8') as f:
                    gpx = gpxpy.parse(f)

                # Extract distance
                desc_text  = gpx.description if gpx.description else ""
                dist_match = re.search(r'Distance: ([\d.]+)km', desc_text)
                distance_str = f"{dist_match.group(1)} km" if dist_match else "Unknown"

                # Distance filter
                min_km = CATEGORY_MIN_KM.get(category)
                if min_km is not None:
                    try:
                        dist_val = float(dist_match.group(1)) if dist_match else 0.0
                        if dist_val < min_km:
                            print(f"{base_name[:35]:<35} | {category:<12} | {CLR_YELLOW}SKIPPED ({dist_val:.1f} km < {min_km} km){CLR_RESET}")
                            continue
                    except ValueError:
                        pass
                print(f"{base_name[:35]:<35} | {category:<12} | {distance_str:<10}")
                # Generate KML
                kml = simplekml.Kml()
                for track in gpx.tracks:
                    ls = kml.newlinestring(name=id_str)
                    ls.coords = [(p.longitude, p.latitude, p.elevation)
                                 for seg in track.segments for p in seg.points]
                    ls.style.linestyle.color = color
                    ls.style.linestyle.width = 4
                    ls.description = (
                        f"<a href=\"https://www.komoot.com/tour/{id_str}\" target=\"_blank\">Open on Komoot</a>"
                    )

                kml.save(kml_path)

            except Exception as e:
                print(f"  {CLR_RED}ERROR: Failed to process {base_name}: {e}{CLR_RESET}")


def sync():
    if not EMAIL or not PASSWORD:
        raise RuntimeError("Set KOMOOT_EMAIL and KOMOOT_PASSWORD environment variables.")
    for cat in set(SPORT_MAP.values()):
        os.makedirs(os.path.join(BASE_DIR, cat), exist_ok=True)
    print(f"{CLR_CYAN}Starting Komoot sync...{CLR_RESET}\n")
    processed = set()  # avoid running same category twice back-to-back
    for sport, category in SPORT_MAP.items():
        out_dir = os.path.join(BASE_DIR, category)
        print(f"  [{category:<10}] sport: {sport}")
        cmd = [
            sys.executable, "-m", "komootgpx",
            f"--mail={EMAIL}",
            f"--pass={PASSWORD}",
            "--make-all",
            "--skip-existing",
            "--id-filename",
            "--tour-type=recorded",
            f"--sport={sport}",
            f"--output={out_dir}",
            "--no-poi",
        ]
        with subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True) as proc:
            for line in proc.stdout:
                line = line.rstrip()
                if line:
                    print(f"{line}")

        print(f"\n{CLR_GREEN}Sync complete.{CLR_RESET}")

def merge_kml_files(base_kml_dir, output_file):
    """Merges all KML files from category subfolders into one master KML file."""
    kml_ns = "http://www.opengis.net/kml/2.2"
    ET.register_namespace("", kml_ns)
    ns = {"kml": kml_ns}
    root     = ET.Element(f"{{{kml_ns}}}kml")
    document = ET.SubElement(root, f"{{{kml_ns}}}Document")
    ET.SubElement(document, f"{{{kml_ns}}}name").text = "Merged Komoot Collection"
    kml_files = glob.glob(os.path.join(base_kml_dir, "**", "*.kml"), recursive=True)
    output_filename = os.path.basename(output_file)
    # Exclude the master file itself in case it ends up in the scan path
    kml_files = [f for f in kml_files if os.path.basename(f) != output_filename]
    if not kml_files:
        print(f"{CLR_YELLOW}No KML files found in '{base_kml_dir}'.{CLR_RESET}")
        return

    print(f"\nMerging {len(kml_files)} KML files into '{output_filename}'...")
    success_count = 0
    error_count   = 0
    for index, kml_path in enumerate(kml_files):
        try:
            tree      = ET.parse(kml_path)
            file_root = tree.getroot()
            prefix    = f"id{index}_"
            for style in file_root.findall(".//kml:Style", ns):
                if 'id' in style.attrib:
                    style.attrib['id'] = prefix + style.attrib['id']
                document.append(style)
            for placemark in file_root.findall(".//kml:Placemark", ns):
                style_url = placemark.find("kml:styleUrl", ns)
                if style_url is not None and style_url.text.startswith("#"):
                    style_url.text = "#" + prefix + style_url.text[1:]
                document.append(placemark)
            success_count += 1
        except Exception as e:
            print(f"  {CLR_RED}ERROR: Could not merge {os.path.basename(kml_path)}: {e}{CLR_RESET}")
            error_count += 1
    try:
        ET.ElementTree(root).write(output_file, encoding="utf-8", xml_declaration=True)
        print("-" * 65)
        print(f"{CLR_GREEN}SUCCESS: {success_count} files merged into '{output_file}'.{CLR_RESET}")
        if error_count:
            print(f"{CLR_RED}ERRORS:  {error_count} files could not be merged.{CLR_RESET}")
    except Exception as e:
        print(f"{CLR_RED}FATAL ERROR: Could not save merged file: {e}{CLR_RESET}")

def ask_sync_or_refresh():
    while True:
        choice = input("Do you want to (S)ync or (R)efresh everything? [S/R]: ").strip().upper()
        if choice in ["S", "R"]:
            return choice
        print("Invalid input. Please enter 'S' or 'R'.")

if __name__ == "__main__":
    choice = ask_sync_or_refresh()
    if choice == "R":
        for folder in ["gpx_files", "kml_files"]:
            if os.path.exists(folder):
                shutil.rmtree(folder)
        print("All folders deleted. Will download everything fresh.")
    check_unknown_sports()
    sync()
    batch_convert_all_categories("gpx_files", "kml_files")
    merge_kml_files("kml_files", "combined.kml")


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
BASE_DIR  = "gpx_files"
SKIP_CACHE = "skipped_ids.txt"

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



def load_skip_cache():
    """Return set of tour IDs known to be below the distance threshold."""
    if not os.path.exists(SKIP_CACHE):
        return set()
    with open(SKIP_CACHE, "r") as f:
        return {line.strip() for line in f if line.strip()}


def save_skip_cache(skipped_ids):
    """Persist the full set of skipped IDs to disk."""
    with open(SKIP_CACHE, "w") as f:
        for id_ in sorted(skipped_ids):
            f.write(id_ + "\n")


def get_activities():
    """Fetch the full list of tour IDs and their sport types from Komoot."""
    cmd = [
        sys.executable, "-m", "komootgpx",
        f"--mail={EMAIL}",
        f"--pass={PASSWORD}",
        "--list-tours",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"ERROR: {result.stderr.strip()}")
        return []

    activities = []
    for line in result.stdout.splitlines():
        line = line.strip()
        if "=>" not in line or "(" not in line:
            continue
        try:
            parts   = line.split(maxsplit=1)
            tour_id = parts[0]
            # Sport type is always the token between '(' and the first ';'
            m = re.search(r'\(([^;)]+);', line)
            if m:
                sport_type = m.group(1).strip().lower()
                activities.append({"id": tour_id, "sport": sport_type})
        except (IndexError, ValueError):
            continue

    return activities


def get_existing_ids(base_dir):
    """
    Return a dict mapping tour ID -> {path, category} for every GPX on disk.
    category is the name of the immediate parent folder (e.g. 'cycling').
    """
    existing = {}
    for gpx_path in glob.glob(os.path.join(base_dir, "**", "*.gpx"), recursive=True):
        stem     = os.path.splitext(os.path.basename(gpx_path))[0]
        id_match = re.search(r"(\d+)", stem)
        if id_match:
            tour_id  = id_match.group(1)
            category = os.path.basename(os.path.dirname(gpx_path))
            existing[tour_id] = {"path": gpx_path, "category": category}
    return existing


def check_and_relocate(activities, base_gpx_dir, base_kml_dir):
    """
    Compare each on-disk GPX's current category folder against what Komoot
    now reports as the sport type.  Move any mismatches to the correct folder
    and delete the stale KML so batch_convert regenerates it in the right place.
    """
    on_disk = get_existing_ids(base_gpx_dir)
    moved   = 0
    unknown = 0

    print(f"\n{CLR_CYAN}Checking sport-type consistency for {len(on_disk)} on-disk files...{CLR_RESET}")

    # Build lookup: tour_id -> correct category according to Komoot
    komoot_category = {}
    for act in activities:
        cat = SPORT_MAP.get(act["sport"])
        if cat:
            komoot_category[act["id"]] = cat

    for tour_id, info in on_disk.items():
        current_cat = info["category"]
        correct_cat = komoot_category.get(tour_id)

        if correct_cat is None:
            # Not in Komoot list or sport unmapped — delete GPX and its KML
            gpx_to_delete = info["path"]
            os.remove(gpx_to_delete)
            old_kml = os.path.join(base_kml_dir, current_cat, tour_id + ".kml")
            if os.path.exists(old_kml):
                os.remove(old_kml)
                kml_note = "(KML deleted)"
            else:
                kml_note = "(no KML found)"
            print(f"  {CLR_RED}DELETED {tour_id} — not in Komoot list or sport unmapped  {kml_note}{CLR_RESET}")
            unknown += 1
            continue

        if current_cat == correct_cat:
            continue  # already in the right place

        # Move GPX to the correct category folder
        src_gpx = info["path"]
        dst_dir = os.path.join(base_gpx_dir, correct_cat)
        dst_gpx = os.path.join(dst_dir, os.path.basename(src_gpx))
        os.makedirs(dst_dir, exist_ok=True)
        shutil.move(src_gpx, dst_gpx)

        # Delete stale KML so it gets regenerated in the new category folder
        old_kml = os.path.join(base_kml_dir, current_cat, tour_id + ".kml")
        if os.path.exists(old_kml):
            os.remove(old_kml)
            kml_note = "(old KML deleted)"
        else:
            kml_note = "(no KML to delete)"

        print(
            f"  MOVED {tour_id}: {CLR_YELLOW}{current_cat}{CLR_RESET}"
            f" -> {CLR_GREEN}{correct_cat}{CLR_RESET}  {kml_note}"
        )
        moved += 1

    # Print summary table: category + sport type -> count
    from collections import defaultdict
    category_sport_map = defaultdict(lambda: defaultdict(int))
    for act in activities:
        cat = SPORT_MAP.get(act["sport"], "unknown")
        category_sport_map[cat][act["sport"]] += 1

    print(f"\n  {'CATEGORY':<12} | {'SPORT TYPE':<18} | {'COUNT':>5}")
    print("  " + "-" * 42)
    total = 0
    for cat in sorted(category_sport_map):
        color = CLR_YELLOW if cat == "unknown" else CLR_GREEN
        for sport in sorted(category_sport_map[cat]):
            count = category_sport_map[cat][sport]
            print(f"  {color}{cat:<12}{CLR_RESET} | {sport:<18} | {count:>5}")
            total += count
    print("  " + "-" * 42)
    print(f"  {'TOTAL':<12} | {'':18} | {total:>5}")

    if moved == 0:
        print(f"\n  {CLR_GREEN}All files are in the correct category folders.{CLR_RESET}")
    else:
        print(f"\n  {CLR_GREEN}Relocated {moved} file(s).{CLR_RESET}")
    if unknown:
        print(f"  {CLR_RED}Deleted  : {unknown} file(s) not in Komoot list or with unmapped sport.{CLR_RESET}")


def download_tour(tour_id, out_dir):
    """Download a single tour by ID into out_dir using komootgpx."""
    cmd = [
        sys.executable, "-m", "komootgpx",
        f"--mail={EMAIL}",
        f"--pass={PASSWORD}",
        "--id-filename",
        "--no-poi",
        f"--make-gpx={tour_id}",
        f"--output={out_dir}",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  {CLR_RED}ERROR downloading {tour_id}: {result.stderr.strip()}{CLR_RESET}")
        return False
    return True


def sync(activities):
    """
    For each activity returned by Komoot, check whether its GPX already exists
    on disk.  Download only the missing ones into the correct category folder.
    """
    if not EMAIL or not PASSWORD:
        raise RuntimeError("Set EMAIL and PASSWORD at the top of the script.")

    for cat in set(SPORT_MAP.values()):
        os.makedirs(os.path.join(BASE_DIR, cat), exist_ok=True)

    on_disk = get_existing_ids(BASE_DIR)
    print(f"\n{CLR_CYAN}Starting Komoot sync...{CLR_RESET}")
    print(f"  {len(activities)} activities on Komoot, {len(on_disk)} already on disk.\n")

    downloaded = 0
    skipped    = 0
    errors     = 0
    unknown    = 0

    for act in activities:
        tour_id  = act["id"]
        sport    = act["sport"]
        category = SPORT_MAP.get(sport)

        if tour_id in on_disk:
            skipped += 1
            continue

        if category is None:
            print(f"  {CLR_YELLOW}UNKNOWN sport '{sport}' for tour {tour_id} — skipping (not downloaded){CLR_RESET}")
            unknown += 1
            continue

        out_dir = os.path.join(BASE_DIR, category)
        print(f"  Downloading tour {tour_id:<12} [{sport} -> {category}] ...", end=" ", flush=True)

        if download_tour(tour_id, out_dir):
            print(f"{CLR_GREEN}OK{CLR_RESET}")
            downloaded += 1
        else:
            errors += 1

    print(f"\n{CLR_GREEN}Sync complete.{CLR_RESET}")
    print(f"  Downloaded : {downloaded}")
    print(f"  Skipped    : {skipped}  (already on disk)")
    if unknown:
        print(f"  Unknown    : {unknown}  (unmapped sport type)")
    if errors:
        print(f"  {CLR_RED}Errors     : {errors}{CLR_RESET}")


def batch_convert_all_categories(base_gpx_dir, base_kml_dir):
    """Convert GPX files from category sub-folders to KML, preserving folder structure."""
    skip_cache = load_skip_cache()
    new_skips  = set()

    print(f"\n{'FILE':<35} | {'CATEGORY':<12} | {'DISTANCE':<10}")
    print("-" * 65)
    for category, color in CATEGORY_COLORS.items():
        gpx_dir = os.path.join(base_gpx_dir, category)
        kml_dir = os.path.join(base_kml_dir, category)
        if not os.path.exists(gpx_dir):
            continue
        os.makedirs(kml_dir, exist_ok=True)
        for gpx_path in glob.glob(os.path.join(gpx_dir, "*.gpx")):
            base_name = os.path.basename(gpx_path)
            stem      = os.path.splitext(base_name)[0]
            id_match  = re.search(r"(\d+)", stem)
            id_str    = id_match.group(1) if id_match else stem
            kml_path  = os.path.join(kml_dir, id_str + ".kml")

            if os.path.exists(kml_path):
                continue

            # Already known to be too short — skip silently
            if id_str in skip_cache:
                continue

            try:
                with open(gpx_path, "r", encoding="utf-8") as f:
                    gpx = gpxpy.parse(f)

                desc_text    = gpx.description or ""
                dist_match   = re.search(r"Distance: ([\d.]+)km", desc_text)
                distance_str = f"{dist_match.group(1)} km" if dist_match else "Unknown"

                min_km = CATEGORY_MIN_KM.get(category)
                if min_km is not None:
                    try:
                        dist_val = float(dist_match.group(1)) if dist_match else 0.0
                        if dist_val < min_km:
                            print(
                                f"{base_name[:35]:<35} | {category:<12} | "
                                f"{CLR_YELLOW}SKIPPED ({dist_val:.1f} km < {min_km} km){CLR_RESET}"
                            )
                            new_skips.add(id_str)
                            continue
                    except ValueError:
                        pass

                print(f"{base_name[:35]:<35} | {category:<12} | {distance_str:<10}")

                kml = simplekml.Kml()
                for track in gpx.tracks:
                    ls        = kml.newlinestring(name=id_str)
                    ls.coords = [
                        (p.longitude, p.latitude, p.elevation)
                        for seg in track.segments
                        for p in seg.points
                    ]
                    ls.style.linestyle.color = color
                    ls.style.linestyle.width = 4
                    ls.description = (
                        f'<a href="https://www.komoot.com/tour/{id_str}" target="_blank">'
                        f"Open on Komoot</a>"
                    )
                kml.save(kml_path)

            except Exception as e:
                print(f"  {CLR_RED}ERROR: Failed to process {base_name}: {e}{CLR_RESET}")

    if new_skips:
        save_skip_cache(skip_cache | new_skips)
        print(f"  {CLR_YELLOW}Cache updated: {len(new_skips)} new ID(s) added to skip cache ({SKIP_CACHE}).{CLR_RESET}")


def merge_kml_files(base_kml_dir, output_file):
    """Merge all KML files from category sub-folders into one master KML file."""
    kml_ns = "http://www.opengis.net/kml/2.2"
    ET.register_namespace("", kml_ns)
    ns = {"kml": kml_ns}

    root     = ET.Element(f"{{{kml_ns}}}kml")
    document = ET.SubElement(root, f"{{{kml_ns}}}Document")
    ET.SubElement(document, f"{{{kml_ns}}}name").text = "Merged Komoot Collection"

    kml_files = glob.glob(os.path.join(base_kml_dir, "**", "*.kml"), recursive=True)
    kml_files = [f for f in kml_files if os.path.basename(f) != os.path.basename(output_file)]

    if not kml_files:
        print(f"{CLR_YELLOW}No KML files found in '{base_kml_dir}'.{CLR_RESET}")
        return

    print(f"\nMerging {len(kml_files)} KML files into '{os.path.basename(output_file)}'...")
    success_count = 0
    error_count   = 0

    for index, kml_path in enumerate(kml_files):
        try:
            file_root = ET.parse(kml_path).getroot()
            prefix    = f"id{index}_"
            for style in file_root.findall(".//kml:Style", ns):
                if "id" in style.attrib:
                    style.attrib["id"] = prefix + style.attrib["id"]
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



if __name__ == "__main__":
    if os.path.exists("credentials.json"):
        os.remove("credentials.json")
        print(f"{CLR_YELLOW}Removed credentials.json{CLR_RESET}")

    # 1. Fetch the authoritative ID + sport list from Komoot
    activities = get_activities()
    if not activities:
        print(f"{CLR_RED}No activities returned — aborting.{CLR_RESET}")
        sys.exit(1)

    # 2. Move any GPX files whose sport type changed to the correct folder
    #    (also deletes stale KMLs so they get regenerated)
    check_and_relocate(activities, "gpx_files", "kml_files")

    # 3. Download activities not yet on disk
    sync(activities)

    # 4. Convert new/relocated GPX files to KML
    batch_convert_all_categories("gpx_files", "kml_files")

    # 5. Rebuild the merged master KML
    merge_kml_files("kml_files", "combined.kml")
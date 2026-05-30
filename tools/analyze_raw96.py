#!/usr/bin/env python3
import argparse
import datetime as _dt
import hashlib
import json
import shutil
import subprocess
from collections import OrderedDict, defaultdict
from pathlib import Path

RAW96_SECTOR_SIZE = 2448
MAIN_SIZE = 2352
SUBCHANNEL_SIZE = 96
SUBQ_OFFSET = MAIN_SIZE + 12
CDTEXT_OFFSET = MAIN_SIZE + 24
LEADIN_SECTORS = 4500
LEADOUT_SECTORS = 6750


def bcd(v):
    return ((v >> 4) * 10) + (v & 0x0F)


def frames_to_msf(frames):
    if frames < 0:
        frames = 0
    mm = frames // (60 * 75)
    ss = (frames % (60 * 75)) // 75
    ff = frames % 75
    return f"{mm:02d}:{ss:02d}:{ff:02d}"


def crc16_ccitt(data):
    crc = 0
    for b in data:
        crc ^= b << 8
        for _ in range(8):
            if crc & 0x8000:
                crc = ((crc << 1) ^ 0x1021) & 0xFFFF
            else:
                crc = (crc << 1) & 0xFFFF
    return crc ^ 0xFFFF


def sha256_file(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def parse_q(q):
    if len(q) != 12:
        return {"crc_ok": False, "raw": q.hex(" ")}
    return {
        "adr_ctl": q[0],
        "track": q[1],
        "index": q[2],
        "rel": f"{bcd(q[3]):02d}:{bcd(q[4]):02d}:{bcd(q[5]):02d}",
        "abs": f"{bcd(q[7]):02d}:{bcd(q[8]):02d}:{bcd(q[9]):02d}",
        "crc_ok": crc16_ccitt(q[:10]) == ((q[10] << 8) | q[11]),
        "raw": q.hex(" "),
    }


def decode_shiftjis_compatible_payload(raw):
    raw = raw.split(b"\x00", 1)[0]
    return raw.decode("cp932", errors="replace")


def detect_sector_size(path):
    if not path.exists():
        raise FileNotFoundError(f"input file not found: {path}")
    size = path.stat().st_size
    candidates = [RAW96_SECTOR_SIZE, MAIN_SIZE, 2368, 2048]
    for sector_size in candidates:
        if size and size % sector_size == 0:
            return sector_size
    return None


def looks_like_cdrom_sector(main):
    return len(main) >= 16 and main[:12] == b"\x00" + (b"\xff" * 10) + b"\x00"


def inspect_program_audio(path, sector_size=MAIN_SIZE, max_sample_sectors=300):
    path = Path(path)
    if not path.exists():
        return {"available": False, "reason": f"missing file: {path}"}

    size = path.stat().st_size
    sectors = size // sector_size if sector_size else None
    profile = {
        "available": True,
        "path": str(path),
        "size_bytes": size,
        "sector_size": sector_size,
        "sectors": sectors,
        "duration": frames_to_msf(sectors or 0),
        "cdda_container_like": sector_size == MAIN_SIZE,
        "sampled_sectors": 0,
        "cdrom_sync_sector_count": 0,
        "all_zero_sector_count": 0,
        "nonzero_byte_count": 0,
        "pcm_sample_min": None,
        "pcm_sample_max": None,
    }
    if sector_size != MAIN_SIZE or not sectors:
        return profile

    sample_limit = min(sectors, max_sample_sectors)
    pcm_min = None
    pcm_max = None
    with path.open("rb") as f:
        for _ in range(sample_limit):
            sec = f.read(sector_size)
            if len(sec) < sector_size:
                break
            profile["sampled_sectors"] += 1
            if looks_like_cdrom_sector(sec):
                profile["cdrom_sync_sector_count"] += 1
            if not any(sec):
                profile["all_zero_sector_count"] += 1
            profile["nonzero_byte_count"] += sum(1 for b in sec if b)
            for i in range(0, len(sec) - 1, 2):
                v = int.from_bytes(sec[i:i + 2], "little", signed=True)
                pcm_min = v if pcm_min is None or v < pcm_min else pcm_min
                pcm_max = v if pcm_max is None or v > pcm_max else pcm_max

    profile["pcm_sample_min"] = pcm_min
    profile["pcm_sample_max"] = pcm_max
    profile["cdda_container_like"] = (
        sector_size == MAIN_SIZE and profile["cdrom_sync_sector_count"] == 0
    )
    profile["has_nonzero_pcm"] = profile["nonzero_byte_count"] > 0
    return profile


def extract_cdtext_from_raw96(data, lead_in):
    cdtext = OrderedDict()
    crc_bad = 0
    for i in range(min(lead_in, len(data) // RAW96_SECTOR_SIZE)):
        sec = data[i * RAW96_SECTOR_SIZE:(i + 1) * RAW96_SECTOR_SIZE]
        rw = sec[CDTEXT_OFFSET:CDTEXT_OFFSET + 72]
        for slot in range(4):
            pack = rw[slot * 18:(slot + 1) * 18]
            if not any(pack):
                continue
            key = (pack[0], pack[1], pack[3] >> 4, pack[3] & 0x0F)
            if crc16_ccitt(pack[:16]) != ((pack[16] << 8) | pack[17]):
                crc_bad += 1
            cdtext.setdefault(key, pack[4:16])

    grouped = defaultdict(list)
    for (ptype, track, block, part), payload in cdtext.items():
        grouped[(ptype, track, block)].append((part, payload))

    texts = []
    for (ptype, track, block), parts in grouped.items():
        kind = {0x80: "TITLE", 0x81: "PERFORMER"}.get(ptype, f"TYPE_{ptype:02X}")
        lang = {0: "EN", 1: "JP"}.get(block, f"BLOCK_{block}")
        raw_payload = b"".join(payload for _, payload in sorted(parts))
        texts.append({
            "kind": kind,
            "track": track,
            "block": block,
            "language": lang,
            "text": decode_shiftjis_compatible_payload(raw_payload),
        })
    return cdtext, crc_bad, texts


def analyze_raw96(path, lead_in=LEADIN_SECTORS, lead_out=LEADOUT_SECTORS):
    if not path.exists():
        raise FileNotFoundError(f"input file not found: {path}")
    data = path.read_bytes()
    if len(data) % RAW96_SECTOR_SIZE:
        raise ValueError(f"size {len(data)} is not divisible by {RAW96_SECTOR_SIZE}")

    total = len(data) // RAW96_SECTOR_SIZE
    program = total - lead_in - lead_out
    if program < 0:
        program = total - lead_in
        lead_out = 0

    samples = []
    sample_points = [
        ("lead-in[0]", 0),
        ("lead-in[1]", 1),
        ("program[0]", lead_in),
        ("program[150]", lead_in + 150),
    ]
    if lead_out:
        sample_points.append(("lead-out[0]", lead_in + program))

    for name, sector_no in sample_points:
        if 0 <= sector_no < total:
            sec = data[sector_no * RAW96_SECTOR_SIZE:(sector_no + 1) * RAW96_SECTOR_SIZE]
            samples.append({
                "name": name,
                "sector": sector_no,
                "q": parse_q(sec[SUBQ_OFFSET:SUBQ_OFFSET + 12]),
                "main_head": sec[:16].hex(" "),
                "main_looks_cdrom": looks_like_cdrom_sector(sec[:MAIN_SIZE]),
            })

    q_bad = 0
    q_sample_limit = min(total, lead_in + min(program, 300) + min(lead_out, 300))
    for i in range(q_sample_limit):
        sec = data[i * RAW96_SECTOR_SIZE:(i + 1) * RAW96_SECTOR_SIZE]
        if not parse_q(sec[SUBQ_OFFSET:SUBQ_OFFSET + 12])["crc_ok"]:
            q_bad += 1

    cdtext, cdtext_crc_bad, texts = extract_cdtext_from_raw96(data, lead_in)
    return {
        "source": str(path),
        "format": "raw96-debug",
        "size_bytes": len(data),
        "sector_size": RAW96_SECTOR_SIZE,
        "total_sectors": total,
        "lead_in_sectors": lead_in,
        "program_sectors": program,
        "lead_out_sectors": lead_out,
        "logical_lba_start": -lead_in,
        "logical_lba_end": program + lead_out - 1,
        "sampled_q_crc_failures": q_bad,
        "cdtext_pack_count": len(cdtext),
        "cdtext_crc_failures": cdtext_crc_bad,
        "cdtext": texts,
        "samples": samples,
    }


def print_raw96_report(meta):
    print(f"file: {meta['source']}")
    print(f"size: {meta['size_bytes']} bytes")
    print(f"sectors: {meta['total_sectors']}")
    print("write mode: DAO RAW+SUB96 semantic test image")
    print(f"lead-in sectors: {meta['lead_in_sectors']}")
    print(f"program sectors: {meta['program_sectors']}")
    print(f"lead-out sectors: {meta['lead_out_sectors']}")
    print(f"logical LBA range: {meta['logical_lba_start']} to {meta['logical_lba_end']}")
    for sample in meta["samples"]:
        print(f"{sample['name']} Q: {sample['q']}")
        print(f"{sample['name']} main[0:16]: {sample['main_head']}")
    print(f"sampled Q CRC failures: {meta['sampled_q_crc_failures']}")
    print(f"unique CD-TEXT packs: {meta['cdtext_pack_count']}")
    print(f"CD-TEXT CRC failures: {meta['cdtext_crc_failures']}")
    for item in meta["cdtext"]:
        print(f"CDTEXT {item['language']} track={item['track']:02d} {item['kind']}: {item['text']}")


def write_json(path, obj):
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def toc_text_from_metadata(meta):
    program = meta["program"]
    tracks = program.get("tracks") or [{
        "number": 1,
        "start_sector": 0,
        "sector_count": program["sectors"],
        "duration": frames_to_msf(program["sectors"]),
    }]
    lines = ["CD_DA", ""]
    cdtext = meta.get("cdtext", [])
    for item in cdtext:
        if item["track"] == 1 and item["kind"] == "TITLE":
            lines += ["// CD-TEXT extracted from RAW+96 lead-in", f'// TRACK 1 TITLE "{item["text"]}"', ""]
    for tr in tracks:
        number = tr["number"]
        start = tr["start_sector"]
        duration = tr["sector_count"]
        lines += [
            f"// Track {number}",
            "TRACK AUDIO",
            "NO COPY",
            "NO PRE_EMPHASIS",
            "TWO_CHANNEL_AUDIO",
            f'FILE "data.bin" {frames_to_msf(start)} {frames_to_msf(duration)}',
            "",
        ]
    return "\n".join(lines) + "\n"


def bundle_from_raw96(raw96_path, out_dir, lead_in=LEADIN_SECTORS, lead_out=LEADOUT_SECTORS, copy_raw96=True):
    raw96_path = raw96_path.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    meta = analyze_raw96(raw96_path, lead_in, lead_out)
    data = raw96_path.read_bytes()
    program_start = meta["lead_in_sectors"]
    program_sectors = meta["program_sectors"]

    data_path = out_dir / "data.bin"
    sub_path = out_dir / "subchannel.bin"
    with data_path.open("wb") as data_out, sub_path.open("wb") as sub_out:
        for i in range(program_sectors):
            sec_no = program_start + i
            sec = data[sec_no * RAW96_SECTOR_SIZE:(sec_no + 1) * RAW96_SECTOR_SIZE]
            data_out.write(sec[:MAIN_SIZE])
            sub_out.write(sec[MAIN_SIZE:MAIN_SIZE + SUBCHANNEL_SIZE])

    if copy_raw96:
        shutil.copy2(raw96_path, out_dir / "raw96-debug.bin")

    metadata = {
        "schema": "kureha-capture-bundle-v1",
        "source_type": "simulated_raw96",
        "created_at": _dt.datetime.now().isoformat(timespec="seconds"),
        "raw96_debug": {
            "path": "raw96-debug.bin" if copy_raw96 else str(raw96_path),
            "sha256": sha256_file(raw96_path),
            "analysis": meta,
        },
        "program": {
            "path": "data.bin",
            "sector_size": MAIN_SIZE,
            "sectors": program_sectors,
            "sha256": sha256_file(data_path),
            "audio_profile": inspect_program_audio(data_path, MAIN_SIZE),
            "tracks": [{
                "number": 1,
                "mode": "AUDIO",
                "start_sector": 0,
                "sector_count": program_sectors,
                "duration": frames_to_msf(program_sectors),
            }],
            "main_samples": [s for s in meta["samples"] if s["name"].startswith("program")],
        },
        "subchannel": {
            "path": "subchannel.bin",
            "available": True,
            "sector_size": SUBCHANNEL_SIZE,
            "sectors": program_sectors,
            "sha256": sha256_file(sub_path),
        },
        "layout": {
            "lead_in_sectors": meta["lead_in_sectors"],
            "program_sectors": program_sectors,
            "lead_out_sectors": meta["lead_out_sectors"],
            "logical_lba_start": meta["logical_lba_start"],
            "logical_lba_end": meta["logical_lba_end"],
        },
        "cdtext": meta["cdtext"],
        "validation": {
            "sampled_q_crc_failures": meta["sampled_q_crc_failures"],
            "cdtext_crc_failures": meta["cdtext_crc_failures"],
            "cdtext_pack_count": meta["cdtext_pack_count"],
        },
    }
    write_json(out_dir / "metadata.json", metadata)
    (out_dir / "disc.toc").write_text(toc_text_from_metadata(metadata), encoding="utf-8")
    return metadata


def read_metadata(bundle_dir):
    meta_path = Path(bundle_dir) / "metadata.json"
    if not meta_path.exists():
        raise FileNotFoundError(f"metadata.json not found in capture bundle: {bundle_dir}")
    return json.loads(meta_path.read_text(encoding="utf-8"))


def analyze_capture(bundle_dir):
    bundle_dir = Path(bundle_dir)
    meta = read_metadata(bundle_dir)
    print(f"bundle: {bundle_dir}")
    print(f"schema: {meta.get('schema')}")
    print(f"source type: {meta.get('source_type')}")
    program = meta.get("program", {})
    layout = meta.get("layout", {})
    print(f"program data: {program.get('path')} ({program.get('sector_size')} bytes/sector, {program.get('sectors')} sectors)")
    audio_profile = program.get("audio_profile")
    if not audio_profile and program.get("path"):
        audio_profile = inspect_program_audio(bundle_dir / program.get("path"), program.get("sector_size") or MAIN_SIZE)
    if audio_profile:
        print(
            "program audio: "
            f"cdda_like={audio_profile.get('cdda_container_like')} "
            f"duration={audio_profile.get('duration')} "
            f"sampled={audio_profile.get('sampled_sectors')} "
            f"cdrom_sync={audio_profile.get('cdrom_sync_sector_count')} "
            f"nonzero_bytes={audio_profile.get('nonzero_byte_count')}"
        )
    print(f"layout: lead-in={layout.get('lead_in_sectors')} program={layout.get('program_sectors')} lead-out={layout.get('lead_out_sectors')}")
    sub = meta.get("subchannel", {})
    print(f"subchannel: {'available' if sub.get('available') else 'unavailable'}")
    validation = meta.get("validation", {})
    if validation:
        print(f"validation: Q CRC failures={validation.get('sampled_q_crc_failures')} CD-TEXT CRC failures={validation.get('cdtext_crc_failures')}")
    for tr in program.get("tracks", []):
        print(f"TRACK {tr.get('number'):02d}: {tr.get('mode')} start={tr.get('start_sector')} sectors={tr.get('sector_count')} duration={tr.get('duration')}")
    for item in meta.get("cdtext", []):
        print(f"CDTEXT {item.get('language')} track={item.get('track'):02d} {item.get('kind')}: {item.get('text')}")


def compare_capture(left_dir, right_dir):
    left = read_metadata(left_dir)
    right = read_metadata(right_dir)
    print(f"left:  {left_dir} ({left.get('source_type')})")
    print(f"right: {right_dir} ({right.get('source_type')})")

    def status(label, a, b):
        marker = "OK" if a == b else "DIFF"
        print(f"[{marker}] {label}: {a!r} vs {b!r}")

    status("program sector size", left.get("program", {}).get("sector_size"), right.get("program", {}).get("sector_size"))
    status("program sectors", left.get("program", {}).get("sectors"), right.get("program", {}).get("sectors"))
    status(
        "program CDDA-like",
        left.get("program", {}).get("audio_profile", {}).get("cdda_container_like"),
        right.get("program", {}).get("audio_profile", {}).get("cdda_container_like"),
    )
    status(
        "program CD-ROM sync sectors",
        left.get("program", {}).get("audio_profile", {}).get("cdrom_sync_sector_count"),
        right.get("program", {}).get("audio_profile", {}).get("cdrom_sync_sector_count"),
    )
    status("track count", len(left.get("program", {}).get("tracks", [])), len(right.get("program", {}).get("tracks", [])))

    left_cdtext = {(x.get("track"), x.get("language"), x.get("kind")): x.get("text") for x in left.get("cdtext", [])}
    right_cdtext = {(x.get("track"), x.get("language"), x.get("kind")): x.get("text") for x in right.get("cdtext", [])}
    for key in sorted(set(left_cdtext) | set(right_cdtext)):
        status(f"CDTEXT {key}", left_cdtext.get(key), right_cdtext.get(key))

    left_sub = left.get("subchannel", {}).get("available")
    right_sub = right.get("subchannel", {}).get("available")
    if left_sub and right_sub:
        status("subchannel sectors", left.get("subchannel", {}).get("sectors"), right.get("subchannel", {}).get("sectors"))
        status("subchannel sha256", left.get("subchannel", {}).get("sha256"), right.get("subchannel", {}).get("sha256"))
    else:
        print(f"[SKIP] subchannel compare: left available={left_sub}, right available={right_sub}")


def run_command(cmd, cwd, out_path, dry_run=False):
    print("+ " + " ".join(cmd))
    if dry_run:
        return
    proc = subprocess.run(cmd, cwd=cwd, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    out_path.write_text(proc.stdout, encoding="utf-8", errors="replace")
    if proc.returncode != 0:
        raise RuntimeError(f"command failed with exit {proc.returncode}: {' '.join(cmd)}\nsee {out_path}")


def capture_disc(out_dir, device=None, name=None, dry_run=False):
    if name:
        out_dir = out_dir / name
    if not dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)
    drutil_prefix = ["drutil"]
    cdrdao_read_cd = ["cdrdao", "read-cd"]
    if device:
        cdrdao_read_cd += ["--device", device]
        drutil_prefix += ["-drive", device]

    commands = [
        (drutil_prefix + ["status"], "drutil-status.txt"),
        (drutil_prefix + ["toc"], "drutil-toc.txt"),
        (drutil_prefix + ["trackinfo"], "drutil-trackinfo.txt"),
        (drutil_prefix + ["cdtext"], "drutil-cdtext.txt"),
        (drutil_prefix + ["subchannel"], "drutil-subchannel.txt"),
        (cdrdao_read_cd + ["--read-raw", "--read-subchan", "rw_raw", "--datafile", "data.bin", "disc.toc"], "cdrdao-read-cd.txt"),
    ]
    for cmd, log_name in commands:
        run_command(cmd, out_dir, out_dir / log_name, dry_run=dry_run)

    if dry_run:
        print(f"dry-run target bundle: {out_dir}")
        return

    data_path = out_dir / "data.bin"
    sector_size = detect_sector_size(data_path) if data_path.exists() else None
    metadata = {
        "schema": "kureha-capture-bundle-v1",
        "source_type": "physical_disc_cdrdao",
        "created_at": _dt.datetime.now().isoformat(timespec="seconds"),
        "capture_commands": [" ".join(cmd) for cmd, _ in commands],
        "program": {
            "path": "data.bin",
            "sector_size": sector_size,
            "sectors": data_path.stat().st_size // sector_size if data_path.exists() and sector_size else None,
            "sha256": sha256_file(data_path) if data_path.exists() else None,
            "audio_profile": inspect_program_audio(data_path, sector_size) if data_path.exists() and sector_size else None,
            "tracks": [],
        },
        "subchannel": {
            "available": False,
            "reason": "cdrdao/drutil outputs are preserved, but no standalone program subchannel.bin was extracted yet",
        },
        "layout": {},
        "cdtext": [],
        "raw_outputs": {
            "toc": "disc.toc",
            "drutil_cdtext": "drutil-cdtext.txt",
            "drutil_toc": "drutil-toc.txt",
            "drutil_trackinfo": "drutil-trackinfo.txt",
            "drutil_subchannel": "drutil-subchannel.txt",
        },
    }
    write_json(out_dir / "metadata.json", metadata)
    print(f"capture bundle written: {out_dir}")


def main():
    parser = argparse.ArgumentParser(description="Inspect and compare Kureha RAW+96 images and capture bundles.")
    sub = parser.add_subparsers(dest="cmd")

    p = sub.add_parser("analyze-bin", help="Analyze a raw96-debug bin or plain data bin.")
    p.add_argument("bin", type=Path)
    p.add_argument("--lead-in", type=int, default=LEADIN_SECTORS)
    p.add_argument("--lead-out", type=int, default=LEADOUT_SECTORS)

    p = sub.add_parser("bundle-from-raw96", help="Export a simulated capture bundle from raw96-debug.bin.")
    p.add_argument("raw96", type=Path)
    p.add_argument("out_dir", type=Path)
    p.add_argument("--lead-in", type=int, default=LEADIN_SECTORS)
    p.add_argument("--lead-out", type=int, default=LEADOUT_SECTORS)
    p.add_argument("--no-copy-raw96", action="store_true")

    p = sub.add_parser("analyze-capture", help="Analyze a capture bundle directory.")
    p.add_argument("bundle_dir", type=Path)

    p = sub.add_parser("compare-capture", help="Compare two capture bundles.")
    p.add_argument("left", type=Path)
    p.add_argument("right", type=Path)

    p = sub.add_parser("capture-disc", help="Capture a physical disc into a bundle using drutil/cdrdao.")
    p.add_argument("out_dir", type=Path)
    p.add_argument("--name")
    p.add_argument("--device")
    p.add_argument("--dry-run", action="store_true")

    args = parser.parse_args()

    if args.cmd is None:
        # Backward compatible one-argument behavior for older test notes.
        if len(__import__("sys").argv) == 2:
            meta = analyze_raw96(Path(__import__("sys").argv[1]), LEADIN_SECTORS, LEADOUT_SECTORS)
            print_raw96_report(meta)
            return 0
        parser.print_help()
        return 2

    if args.cmd == "analyze-bin":
        sector_size = detect_sector_size(args.bin)
        if sector_size == RAW96_SECTOR_SIZE:
            print_raw96_report(analyze_raw96(args.bin, args.lead_in, args.lead_out))
        else:
            size = args.bin.stat().st_size
            print(f"file: {args.bin}")
            print(f"size: {size} bytes")
            print(f"sector size: {sector_size or 'unknown'}")
            print(f"sectors: {size // sector_size if sector_size else 'unknown'}")
            if sector_size == MAIN_SIZE:
                audio_profile = inspect_program_audio(args.bin, MAIN_SIZE)
                print(
                    "program audio: "
                    f"cdda_like={audio_profile.get('cdda_container_like')} "
                    f"duration={audio_profile.get('duration')} "
                    f"sampled={audio_profile.get('sampled_sectors')} "
                    f"cdrom_sync={audio_profile.get('cdrom_sync_sector_count')} "
                    f"nonzero_bytes={audio_profile.get('nonzero_byte_count')}"
                )
        return 0
    if args.cmd == "bundle-from-raw96":
        metadata = bundle_from_raw96(args.raw96, args.out_dir, args.lead_in, args.lead_out, not args.no_copy_raw96)
        print(f"capture bundle written: {args.out_dir}")
        print(f"program sectors: {metadata['program']['sectors']}")
        print(f"CD-TEXT entries: {len(metadata['cdtext'])}")
        return 0
    if args.cmd == "analyze-capture":
        analyze_capture(args.bundle_dir)
        return 0
    if args.cmd == "compare-capture":
        compare_capture(args.left, args.right)
        return 0
    if args.cmd == "capture-disc":
        capture_disc(args.out_dir, args.device, args.name, args.dry_run)
        return 0
    return 2


if __name__ == "__main__":
    raise SystemExit(main())

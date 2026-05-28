#!/usr/bin/env python3
import argparse
from collections import OrderedDict, defaultdict
from pathlib import Path

SECTOR_SIZE = 2448
MAIN_SIZE = 2352
SUBQ_OFFSET = MAIN_SIZE + 12
CDTEXT_OFFSET = MAIN_SIZE + 24
LEADIN_SECTORS = 4500
LEADOUT_SECTORS = 6750


def bcd(v):
    return ((v >> 4) * 10) + (v & 0x0F)


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


def parse_q(q):
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
    # VB6 on Japanese Windows emits CP932 bytes; this is the practical
    # Shift-JIS-compatible form we need to inspect for CD-TEXT.
    return raw.decode("cp932", errors="replace")


def main():
    ap = argparse.ArgumentParser(description="Inspect Kureha file-backed RAW+96 test images.")
    ap.add_argument("bin", type=Path)
    ap.add_argument("--lead-in", type=int, default=LEADIN_SECTORS)
    ap.add_argument("--lead-out", type=int, default=LEADOUT_SECTORS)
    args = ap.parse_args()

    data = args.bin.read_bytes()
    if len(data) % SECTOR_SIZE:
        print(f"ERROR: size {len(data)} is not divisible by {SECTOR_SIZE}")
        return 2

    total = len(data) // SECTOR_SIZE
    program = total - args.lead_in - args.lead_out
    if program < 0:
        program = total - args.lead_in
        args.lead_out = 0
    print(f"file: {args.bin}")
    print(f"size: {len(data)} bytes")
    print(f"sectors: {total}")
    print("write mode: DAO RAW+SUB96 semantic test image")
    print(f"lead-in sectors: {args.lead_in}")
    print(f"program sectors: {program}")
    print(f"lead-out sectors: {args.lead_out}")
    print(f"logical LBA range: {-args.lead_in} to {program + args.lead_out - 1}")

    sample_points = [
        ("lead-in[0]", 0),
        ("lead-in[1]", 1),
        ("program[0]", args.lead_in),
        ("program[150]", args.lead_in + 150),
    ]
    if args.lead_out:
        sample_points.append(("lead-out[0]", args.lead_in + program))

    for name, sector_no in sample_points:
        if 0 <= sector_no < total:
            sec = data[sector_no * SECTOR_SIZE:(sector_no + 1) * SECTOR_SIZE]
            print(f"{name} Q: {parse_q(sec[SUBQ_OFFSET:SUBQ_OFFSET + 12])}")
            print(f"{name} main[0:16]: {sec[:16].hex(' ')}")

    q_bad = 0
    q_sample_limit = min(total, args.lead_in + min(program, 300) + min(args.lead_out, 300))
    for i in range(q_sample_limit):
        sec = data[i * SECTOR_SIZE:(i + 1) * SECTOR_SIZE]
        if not parse_q(sec[SUBQ_OFFSET:SUBQ_OFFSET + 12])["crc_ok"]:
            q_bad += 1
    print(f"sampled Q CRC failures: {q_bad}")

    cdtext = OrderedDict()
    cdtext_crc_bad = 0
    for i in range(min(args.lead_in, total)):
        sec = data[i * SECTOR_SIZE:(i + 1) * SECTOR_SIZE]
        rw = sec[CDTEXT_OFFSET:CDTEXT_OFFSET + 72]
        for slot in range(4):
            pack = rw[slot * 18:(slot + 1) * 18]
            if not any(pack):
                continue
            key = (pack[0], pack[1], pack[3] >> 4, pack[3] & 0x0F)
            if crc16_ccitt(pack[:16]) != ((pack[16] << 8) | pack[17]):
                cdtext_crc_bad += 1
            cdtext.setdefault(key, pack[4:16])

    print(f"unique CD-TEXT packs: {len(cdtext)}")
    print(f"CD-TEXT CRC failures: {cdtext_crc_bad}")
    grouped = defaultdict(list)
    for (ptype, track, block, part), payload in cdtext.items():
        grouped[(ptype, track, block)].append((part, payload))

    for (ptype, track, block), parts in grouped.items():
        kind = {0x80: "TITLE", 0x81: "PERFORMER"}.get(ptype, f"TYPE_{ptype:02X}")
        lang = {0: "EN", 1: "JP"}.get(block, f"BLOCK_{block}")
        raw_payload = b"".join(payload for _, payload in sorted(parts))
        print(f"CDTEXT {lang} track={track:02d} {kind}: {decode_shiftjis_compatible_payload(raw_payload)}")


if __name__ == "__main__":
    raise SystemExit(main())

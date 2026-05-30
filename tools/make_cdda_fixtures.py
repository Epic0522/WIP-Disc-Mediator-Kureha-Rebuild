#!/usr/bin/env python3
import argparse
import math
import wave
from pathlib import Path

SAMPLE_RATE = 44100
CHANNELS = 2
SAMPLE_WIDTH = 2


def write_stereo_pcm16(path, seconds, sample_fn):
    total_frames = int(SAMPLE_RATE * seconds)
    path.parent.mkdir(parents=True, exist_ok=True)
    with wave.open(str(path), "wb") as wav:
        wav.setnchannels(CHANNELS)
        wav.setsampwidth(SAMPLE_WIDTH)
        wav.setframerate(SAMPLE_RATE)
        frames = bytearray()
        for i in range(total_frames):
            value = int(max(-1.0, min(1.0, sample_fn(i))) * 32767)
            sample = value.to_bytes(2, "little", signed=True)
            frames += sample
            frames += sample
        wav.writeframes(bytes(frames))


def main():
    parser = argparse.ArgumentParser(description="Create deterministic 44.1 kHz / 16-bit stereo WAV fixtures for Kureha CDDA tests.")
    parser.add_argument("out_dir", nargs="?", default="captures/fixtures", type=Path)
    args = parser.parse_args()

    out_dir = args.out_dir
    write_stereo_pcm16(out_dir / "cdda_silence_2s.wav", 2.0, lambda _i: 0.0)
    write_stereo_pcm16(out_dir / "cdda_sine_440hz_2s.wav", 2.0, lambda i: 0.35 * math.sin((2.0 * math.pi * 440.0 * i) / SAMPLE_RATE))
    write_stereo_pcm16(out_dir / "アンティーカ_テスト_2s.wav", 2.0, lambda i: 0.25 * math.sin((2.0 * math.pi * 880.0 * i) / SAMPLE_RATE))

    print(f"fixtures written: {out_dir}")
    for path in sorted(out_dir.glob("*.wav")):
        print(path)


if __name__ == "__main__":
    raise SystemExit(main())

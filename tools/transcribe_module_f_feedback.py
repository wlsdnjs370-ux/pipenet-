"""Create a timestamped Korean transcript for the supplied feedback recording."""

import os
import sysconfig
from pathlib import Path

_site_packages = Path(sysconfig.get_paths()["purelib"])
os.environ["PATH"] = os.pathsep.join(
    [
        str(_site_packages / "nvidia" / "cublas" / "bin"),
        str(_site_packages / "nvidia" / "cudnn" / "bin"),
        os.environ["PATH"],
    ]
)

from faster_whisper import WhisperModel


AUDIO = Path(r"C:/Users/admin/PycharmProjects/JupyterProject/outputs/module_f_feedback_928_focus_enhanced.wav")
OUTPUT = Path(r"C:/Users/admin/PycharmProjects/JupyterProject/outputs/module_f_feedback_928_early_transcript.txt")
TIMESTAMP_OFFSET = 1000.0


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    model = WhisperModel("large-v3", device="cuda", compute_type="float16")
    segments, info = model.transcribe(
        str(AUDIO),
        language="ko",
        beam_size=5,
        vad_filter=True,
        vad_parameters={"min_silence_duration_ms": 450},
        condition_on_previous_text=False,
        clip_timestamps="0,610",
        hallucination_silence_threshold=2.0,
    )
    with OUTPUT.open("w", encoding="utf-8") as transcript:
        transcript.write(f"언어: {info.language} (신뢰도 {info.language_probability:.2f})\n\n")
        for segment in segments:
            transcript.write(
                f"[{segment.start + TIMESTAMP_OFFSET:07.2f}–{segment.end + TIMESTAMP_OFFSET:07.2f}] {segment.text.strip()}\n"
            )


if __name__ == "__main__":
    main()

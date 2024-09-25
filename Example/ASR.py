from funasr import AutoModel
from funasr.utils.postprocess_utils import rich_transcription_postprocess

model_dir = "iic/SenseVoiceSmall"

model = AutoModel(
    model=model_dir,
    vad_model="fsmn-vad",
    vad_kwargs={"max_single_segment_time": 30000},
    device="cuda:0",
)

# en
import tkinter as tk
from tkinter import filedialog
import os

root = tk.Tk()
root.withdraw()

file_paths = filedialog.askopenfilenames(title="选择音频文件", filetypes=[("音频文件", "*.mp3 *.wav")])

for file_path in file_paths:
    res = model.generate(
        input=file_path,
        cache={},
        language="auto",  # "zn", "en", "yue", "ja", "ko", "nospeech"
        use_itn=True,
        batch_size_s=60,
        merge_vad=True,  #
        merge_length_s=15,
    )
    text = rich_transcription_postprocess(res[0]["text"])
    
    # 保存文本到同名txt文件
    output_path = os.path.splitext(file_path)[0] + ".txt"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)
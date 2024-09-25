from flask import Flask, request, jsonify
import os
from funasr import AutoModel
from funasr.utils.postprocess_utils import rich_transcription_postprocess

MODEL_DIR = "iic/SenseVoiceSmall"
DEVICE = "cuda:0"
TMP_DIR = "/tmp"


class ASRService:
    def __init__(self, model_dir, device, tmp_dir, port=5000):
        self.model_dir = model_dir
        self.device = device
        self.tmp_dir = tmp_dir
        self.port = port
        self.model = self._initialize_model()
        self.app = Flask(__name__)
        self._configure_app()
        self._setup_routes()

    def _initialize_model(self):
        return AutoModel(
            model=self.model_dir,
            vad_model="fsmn-vad",
            vad_kwargs={"max_single_segment_time": 30000},
            device=self.device,
            disable_update=True,
        )

    def _configure_app(self):
        self.app.config.update(
            SERVICE_NAME='ASR Service',
            SERVICE_DESCRIPTION='A service for converting speech to text using ASR models.'
        )

    def _save_file(self, file):
        file_path = os.path.join(self.tmp_dir, file.filename)
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        file.save(file_path)
        return file_path

    def _generate_text(self, file_path):
        res = self.model.generate(
            input=file_path,
            cache={},
            language="auto",
            use_itn=True,
            batch_size_s=60,
            merge_vad=True,
            merge_length_s=15,
        )
        return rich_transcription_postprocess(res[0]["text"])

    def _setup_routes(self):
        @self.app.route('/process_audio', methods=['POST'])
        def process_audio():
            if 'file' not in request.files or request.files['file'].filename == '':
                return jsonify({"error": "No file part or no selected file"}), 400

            file_path = self._save_file(request.files['file'])
            try:
                text = self._generate_text(file_path)
            finally:
                os.remove(file_path)

            return jsonify({"text": text})

        @self.app.route('/check_connection', methods=['GET'])
        def check_connection():
            return jsonify({"status": "success", "message": "Connection successful"}), 200

    def start_server(self):
        self.app.run(host='0.0.0.0', port=self.port)

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="ASR Service")
    parser.add_argument('--port', type=int, default=5000, help='Port to run the server on')
    args = parser.parse_args()

    asr_service = ASRService(MODEL_DIR, DEVICE, TMP_DIR, port=args.port)
    asr_service.start_server()
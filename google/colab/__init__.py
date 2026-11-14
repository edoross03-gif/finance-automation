# Mock of google.colab module for local dashboard execution

class FilesMock:
    def __init__(self):
        self._upload_count = 0

    def upload(self):
        import sys
        import os
        
        self._upload_count += 1
        print(f"📥 [Mock Colab] files.upload() called (Upload #{self._upload_count})")
        
        # Check environment variable INPUT_FILES (comma-separated list of local files)
        input_files_str = os.environ.get("INPUT_FILES", "")
        if input_files_str:
            input_files = [f.strip() for f in input_files_str.split(",") if f.strip()]
        else:
            input_files = []
            
        # Also check sys.argv for file paths
        for arg in sys.argv[1:]:
            if "=" not in arg and os.path.exists(arg) and os.path.isfile(arg):
                if arg not in input_files:
                    input_files.append(arg)
                    
        # Fallback to files in uploads/ if nothing specified
        if not input_files:
            uploads_dir = "uploads"
            if os.path.exists(uploads_dir):
                all_files = [os.path.join(uploads_dir, f) for f in os.listdir(uploads_dir) if os.path.isfile(os.path.join(uploads_dir, f))]
                if all_files:
                    # Sort by modification time to get the most recent ones first
                    all_files.sort(key=os.path.getmtime, reverse=True)
                    input_files = all_files

        if not input_files:
            print("⚠️ [Mock Colab] No local files found for upload simulation.")
            return {}

        # Select the file corresponding to the current upload count
        # (index is self._upload_count - 1, wrapped using modulo if we have fewer files than calls)
        idx = (self._upload_count - 1) % len(input_files)
        selected_file = input_files[idx]
        
        filename = os.path.basename(selected_file)
        print(f"   -> Simulating upload of file: '{selected_file}'")
        try:
            with open(selected_file, 'rb') as f:
                content = f.read()
                
            # Scrivi fisicamente il file nella directory corrente dell'esecuzione (simulando il comportamento di Colab)
            with open(filename, 'wb') as out_f:
                out_f.write(content)
            print(f"   -> [Mock Colab] Saved local copy to current directory: '{filename}'")
            
            return {filename: content}
        except Exception as e:
            print(f"❌ [Mock Colab] Error reading/writing file '{selected_file}': {e}")
            return {}

    def download(self, filename):
        import os
        import shutil
        print(f"📤 [Mock Colab] files.download('{filename}') called")
        
        # Verify if the output file exists
        if not os.path.exists(filename):
            print(f"⚠️ [Mock Colab] Warning: Output file '{filename}' was not found in current directory.")
            return
            
        # Copy the downloaded file to the uploads/ directory so the dashboard can offer it for download!
        uploads_dir = "uploads"
        if os.path.exists(uploads_dir):
            dest = os.path.join(uploads_dir, os.path.basename(filename))
            try:
                shutil.copy(filename, dest)
                print(f"💾 [Mock Colab] Saved output copy to: {dest}")
            except Exception as e:
                print(f"❌ [Mock Colab] Error saving output copy to '{dest}': {e}")

files = FilesMock()

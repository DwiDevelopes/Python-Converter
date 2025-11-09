import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
from threading import Thread
import sys
import tarfile
import platform

class AutoFFmpegConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto FFmpeg Video Converter")
        self.root.geometry("700x500")
        
        # Variabel FFmpeg
        self.ffmpeg_path = None
        self.ffmpeg_extracted = False
        
        # Setup UI
        self.setup_ui()
        
        # Cek FFmpeg
        self.check_ffmpeg()
    
    def setup_ui(self):
        # Style
        self.style = ttk.Style()
        self.style.configure('TButton', padding=5)
        self.style.configure('TLabel', padding=5)
        
        # Variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.quality_var = tk.StringVar(value="balanced")
        self.status_var = tk.StringVar(value="Ready")
        
        # Main Frame
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # FFmpeg Section
        ffmpeg_frame = ttk.LabelFrame(main_frame, text="FFmpeg Setup", padding="10")
        ffmpeg_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(ffmpeg_frame, text="FFmpeg Archive (ffmpeg-7.1.1.tar.xz):").grid(row=0, column=0, sticky=tk.W)
        self.ffmpeg_entry = ttk.Entry(ffmpeg_frame, width=50)
        self.ffmpeg_entry.grid(row=0, column=1, padx=5)
        ttk.Button(ffmpeg_frame, text="Browse", command=self.browse_ffmpeg).grid(row=0, column=2)
        ttk.Button(ffmpeg_frame, text="Extract FFmpeg", command=self.extract_ffmpeg).grid(row=1, column=1, pady=5)
        
        # Input Section
        input_frame = ttk.LabelFrame(main_frame, text="Video Conversion", padding="10")
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="Input Video:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_input).grid(row=0, column=2)
        
        ttk.Label(input_frame, text="Output Video:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_output).grid(row=1, column=2)
        
        ttk.Label(input_frame, text="Quality Preset:").grid(row=2, column=0, sticky=tk.W)
        ttk.Combobox(input_frame, textvariable=self.quality_var, 
                    values=["high", "balanced", "medium", "small"], 
                    state="readonly").grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Progress
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=550, mode='determinate')
        self.progress.pack(pady=15)
        
        # Status
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.pack()
        
        # Convert Button
        self.convert_btn = ttk.Button(main_frame, text="Convert Video", command=self.start_conversion, state=tk.DISABLED)
        self.convert_btn.pack(pady=10)
    
    def browse_ffmpeg(self):
        file_path = filedialog.askopenfilename(
            title="Select FFmpeg Archive",
            filetypes=(("XZ Archive", "*.tar.xz"), ("All Files", "*.*"))
        )
        if file_path:
            self.ffmpeg_entry.delete(0, tk.END)
            self.ffmpeg_entry.insert(0, file_path)
    
    def extract_ffmpeg(self):
        archive_path = self.ffmpeg_entry.get()
        if not archive_path:
            messagebox.showerror("Error", "Please select FFmpeg archive file")
            return
        
        try:
            self.status_var.set("Extracting FFmpeg...")
            self.root.update()
            
            # Buat folder ffmpeg jika belum ada
            ffmpeg_dir = os.path.join(os.getcwd(), "ffmpeg")
            os.makedirs(ffmpeg_dir, exist_ok=True)
            
            # Ekstrak archive
            with tarfile.open(archive_path, 'r:xz') as tar:
                tar.extractall(ffmpeg_dir)
            
            # Cari binary FFmpeg
            for root, dirs, files in os.walk(ffmpeg_dir):
                if 'ffmpeg' in files:
                    self.ffmpeg_path = os.path.join(root, 'ffmpeg')
                    break
            
            if platform.system() == "Windows":
                # Untuk Windows, cari ffmpeg.exe
                for root, dirs, files in os.walk(ffmpeg_dir):
                    if 'ffmpeg.exe' in files:
                        self.ffmpeg_path = os.path.join(root, 'ffmpeg.exe')
                        break
            
            if self.ffmpeg_path:
                self.ffmpeg_extracted = True
                self.convert_btn.config(state=tk.NORMAL)
                self.status_var.set("FFmpeg ready to use!")
                messagebox.showinfo("Success", "FFmpeg extracted successfully!")
            else:
                self.status_var.set("FFmpeg binary not found in archive")
                messagebox.showerror("Error", "Could not find FFmpeg binary in the archive")
        
        except Exception as e:
            self.status_var.set("Extraction failed")
            messagebox.showerror("Error", f"Failed to extract FFmpeg: {str(e)}")
    
    def check_ffmpeg(self):
        """Cek jika FFmpeg sudah ada di folder"""
        ffmpeg_dir = os.path.join(os.getcwd(), "ffmpeg")
        
        if platform.system() == "Windows":
            ffmpeg_exe = os.path.join(ffmpeg_dir, "ffmpeg.exe")
            if os.path.exists(ffmpeg_exe):
                self.ffmpeg_path = ffmpeg_exe
                self.ffmpeg_extracted = True
                self.convert_btn.config(state=tk.NORMAL)
                self.status_var.set("FFmpeg ready to use!")
        else:
            ffmpeg_bin = os.path.join(ffmpeg_dir, "ffmpeg")
            if os.path.exists(ffmpeg_bin):
                self.ffmpeg_path = ffmpeg_bin
                self.ffmpeg_extracted = True
                self.convert_btn.config(state=tk.NORMAL)
                self.status_var.set("FFmpeg ready to use!")
    
    def browse_input(self):
        file_path = filedialog.askopenfilename(
            title="Select Video File",
            filetypes=(("Video Files", "*.mp4 *.avi *.mov *.mkv *.flv *.webm"), 
                      ("All Files", "*.*"))
        )
        if file_path:
            self.input_path.set(file_path)
            self.auto_set_output_path(file_path)
    
    def browse_output(self):
        initial_file = self.output_path.get() or os.path.join(os.getcwd(), "output.mp4")
        file_path = filedialog.asksaveasfilename(
            title="Save Output Video",
            defaultextension=".mp4",
            initialfile=os.path.basename(initial_file),
            filetypes=(("MP4 Files", "*.mp4"), ("All Files", "*.*"))
        )
        if file_path:
            self.output_path.set(file_path)
    
    def auto_set_output_path(self, input_path):
        """Set output path automatically"""
        dir_name, file_name = os.path.split(input_path)
        name, ext = os.path.splitext(file_name)
        output_file = f"{name}_compressed.mp4"
        self.output_path.set(os.path.join(dir_name, output_file))
    
    def get_quality_settings(self):
        """Return FFmpeg parameters based on quality selection"""
        quality = self.quality_var.get()
        
        if quality == "high":
            return {"crf": 18, "preset": "slow", "audio_quality": 192}
        elif quality == "balanced":
            return {"crf": 22, "preset": "medium", "audio_quality": 160}
        elif quality == "medium":
            return {"crf": 26, "preset": "fast", "audio_quality": 128}
        else:  # small
            return {"crf": 30, "preset": "veryfast", "audio_quality": 96}
    
    def start_conversion(self):
        if not self.ffmpeg_extracted:
            messagebox.showerror("Error", "Please extract FFmpeg first")
            return
            
        if not self.input_path.get():
            messagebox.showerror("Error", "Please select input video file")
            return
            
        if not self.output_path.get():
            messagebox.showerror("Error", "Please select output video file")
            return
            
        # Start conversion in thread
        Thread(target=self.run_conversion, daemon=True).start()
    
    def run_conversion(self):
        try:
            input_file = self.input_path.get()
            output_file = self.output_path.get()
            
            # Validasi file input
            if not os.path.exists(input_file):
                messagebox.showerror("Error", "Input file does not exist!")
                return
                
            # Cek ekstensi output
            if not output_file.lower().endswith('.mp4'):
                output_file += '.mp4'
                self.output_path.set(output_file)
            
            # Get quality settings
            settings = self.get_quality_settings()
            
            # Update UI
            self.status_var.set("Preparing conversion...")
            self.progress['value'] = 0
            self.root.update()
            
            # Bangun perintah FFmpeg
            cmd = [
                self.ffmpeg_path,
                '-i', input_file,
                '-c:v', 'libx264',
                '-crf', str(settings['crf']),
                '-preset', settings['preset'],
                '-c:a', 'aac',
                '-b:a', f"{settings['audio_quality']}k",
                '-movflags', '+faststart',
                '-y',  # Overwrite output
                output_file
            ]
            
            # Untuk Windows, kita perlu mengatur startupinfo
            startupinfo = None
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            # Jalankan FFmpeg
            process = subprocess.Popen(
                cmd, 
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1,
                startupinfo=startupinfo
            )
            
            # Parse output untuk progress
            duration = 0
            for line in process.stdout:
                line = line.strip()
                
                # Cari durasi video
                if 'Duration:' in line:
                    time_str = line.split('Duration:')[1].split(',')[0].strip()
                    h, m, s = time_str.split(':')
                    duration = int(h)*3600 + int(m)*60 + float(s)
                
                # Update progress
                if 'time=' in line:
                    time_str = line.split('time=')[1].split()[0]
                    h, m, s = time_str.split(':')
                    current_time = int(h)*3600 + int(m)*60 + float(s)
                    
                    if duration > 0:
                        progress_percent = (current_time / duration) * 100
                        self.progress['value'] = progress_percent
                        self.status_var.set(f"Converting... {progress_percent:.1f}%")
                        self.root.update()
            
            process.wait()
            
            if process.returncode == 0:
                self.progress['value'] = 100
                self.status_var.set("Conversion complete!")
                
                # Tampilkan info ukuran file
                input_size = os.path.getsize(input_file) / (1024 * 1024)  # MB
                output_size = os.path.getsize(output_file) / (1024 * 1024)  # MB
                
                messagebox.showinfo(
                    "Success", 
                    f"Video converted successfully!\n\n"
                    f"Original size: {input_size:.2f} MB\n"
                    f"Compressed size: {output_size:.2f} MB\n"
                    f"Saved to: {output_file}"
                )
            else:
                messagebox.showerror("Error", "Video conversion failed")
                self.status_var.set("Conversion failed")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set(f"Error: {str(e)}")
        finally:
            if self.progress['value'] < 100:
                self.progress['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoFFmpegConverter(root)
    root.mainloop()
import speech_recognition as sr
import subprocess
import psutil


class VoiceAppController:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.listening = False

        # Adjust for ambient noise
        print("Adjusting for ambient noise... Please wait.")
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source)
        print("Ready for voice commands!")

    def listen_for_command(self) -> str:
        """Listen for voice command and return the recognized text"""
        try:
            with self.microphone as source:
                print("Listening for command...")
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=3)

            command = self.recognizer.recognize_google(audio).lower()
            print(f"Recognized: {command}")
            return command

        except sr.UnknownValueError:
            print("Could not understand the command")
            return ""
        except sr.RequestError as e:
            print(f"Error with speech recognition service: {e}")
            return ""
        except sr.WaitTimeoutError:
            print("Listening timeout")
            return ""

    def open_application(self, app_name: str) -> bool:
        """Open application using UWP approach first, then fallback methods"""
        app_name = app_name.strip()
        
        # Method 1: Try PowerShell Get-StartApps for UWP apps (primary method)
        try:
            print(f"Searching for '{app_name}' in installed apps...")
            result = subprocess.run(
                ['powershell', '-Command', f'Get-StartApps | Where-Object {{$_.Name -like \"*{app_name}*\"}} | Select-Object -First 1 -ExpandProperty AppID'],
                shell=True, capture_output=True, text=True
            )
            if result.stdout.strip():
                app_id = result.stdout.strip()
                subprocess.run(['start', 'shell:appsFolder\\' + app_id], shell=True, check=True)
                print(f"Opened {app_name} (UWP app)")
                return True
        except subprocess.CalledProcessError:
            pass
        
        # Method 2: Try direct start command (fallback for traditional apps)
        try:
            subprocess.run(['start', app_name], shell=True, check=True)
            print(f"Opened {app_name}")
            return True
        except subprocess.CalledProcessError:
            pass
        
        # Method 3: Try with .exe extension (fallback)
        try:
            subprocess.run(['start', f'{app_name}.exe'], shell=True, check=True)
            print(f"Opened {app_name}.exe")
            return True
        except subprocess.CalledProcessError:
            pass
        
        # If all methods fail, provide detailed error message
        print(f"Could not start application: {app_name}")
        print("Possible reasons:")
        print(f"  - App '{app_name}' is not installed")
        print(f"  - Try saying the exact app name as it appears in Start Menu")
        print(f"  - For Microsoft Store apps, try the exact name from the store")
        print(f"  - Some apps may have different executable names")
        return False

    def close_application(self, app_name: str) -> bool:
        """Close an application by name using multiple methods"""
        app_name = app_name.strip()
        closed_count = 0
        
        try:
            # Method 1: Look for exact .exe match
            executable = f"{app_name}.exe"
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] and executable.lower() == proc.info['name'].lower():
                    try:
                        proc.terminate()
                        closed_count += 1
                        print(f"Closed {proc.info['name']} (PID: {proc.info['pid']})")
                    except psutil.NoSuchProcess:
                        pass
                    except psutil.AccessDenied:
                        print(f"Access denied to close {proc.info['name']}")
            
            # Method 2: Look for partial matches (for apps with different process names)
            if closed_count == 0:
                for proc in psutil.process_iter(['pid', 'name']):
                    if proc.info['name'] and app_name.lower() in proc.info['name'].lower():
                        try:
                            proc.terminate()
                            closed_count += 1
                            print(f"Closed {proc.info['name']} (PID: {proc.info['pid']})")
                        except psutil.NoSuchProcess:
                            pass
                        except psutil.AccessDenied:
                            print(f"Access denied to close {proc.info['name']}")
            
            # Method 3: Try to close UWP apps using PowerShell (enhanced UWP support)
            if closed_count == 0:
                try:
                    # First try to find and close by process name pattern
                    result = subprocess.run(
                        ['powershell', '-Command', f'Get-Process | Where-Object {{$_.ProcessName -like \"*{app_name}*\"}} | Stop-Process -Force'],
                        shell=True, capture_output=True, text=True
                    )
                    if result.returncode == 0:
                        print(f"Attempted to close {app_name} using PowerShell")
                        closed_count = 1  # Assume success if no error
                    
                    # Also try to close UWP apps by finding their main window
                    if closed_count == 0:
                        result = subprocess.run(
                            ['powershell', '-Command', f'Get-Process | Where-Object {{$_.MainWindowTitle -like \"*{app_name}*\"}} | Stop-Process -Force'],
                            shell=True, capture_output=True, text=True
                        )
                        if result.returncode == 0:
                            print(f"Closed {app_name} by window title")
                            closed_count = 1
                except subprocess.CalledProcessError:
                    pass
            
            if closed_count > 0:
                print(f"Successfully closed {app_name}")
                return True
            else:
                print(f"No running instances of '{app_name}' found")
                print("Possible reasons:")
                print(f"  - App '{app_name}' is not currently running")
                print(f"  - App may have a different process name")
                print(f"  - Try saying the exact process name")
                return False

        except Exception as e:
            print(f"Error closing {app_name}: {e}")
            return False

    def process_command(self, command: str):
        """Process the voice command"""
        command = command.lower().strip()

        if "stop listening" in command or "exit" in command:
            self.listening = False
            print("Stopping voice control...")
            return

        if "open" in command:
            app_name = command.replace("open", "").strip()
            if app_name:
                self.open_application(app_name)
            else:
                print("Please specify which application to open")

        elif "close" in command:
            app_name = command.replace("close", "").strip()
            if app_name:
                self.close_application(app_name)
            else:
                print("Please specify which application to close")

        else:
            print("Available commands:")
            print("  - 'open [app name]' - Opens an application")
            print("  - 'close [app name]' - Closes an application")
            print("  - 'stop listening' - Stops the voice controller")

    def start_listening(self):
        """Start the main listening loop"""
        self.listening = True
        print("\n=== Voice App Controller Started ===")
        print("Say commands like:")
        print("  - 'open notepad'")
        print("  - 'close chrome'")
        print("  - 'stop listening' to exit")
        print("=====================================\n")

        while self.listening:
            try:
                command = self.listen_for_command()
                if command:
                    self.process_command(command)
                time.sleep(0.5)  # Brief pause between commands
            except KeyboardInterrupt:
                print("\nStopping voice control...")
                break

        print("Voice controller stopped.")

def main():
    """Main function to run the voice app controller"""
    try:
        controller = VoiceAppController()
        controller.start_listening()
    except Exception as e:
        print(f"Error initializing voice controller: {e}")
        print("Make sure you have the required packages installed:")
        print("pip install speechrecognition pyaudio psutil")

if __name__ == "__main__":
    main()

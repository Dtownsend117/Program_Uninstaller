import subprocess
import winreg
import speech_recognition as sr

def list_installed_programs():
    """List all installed programs on the Windows system."""
    programs = []
    uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" # Windows registry key to uninstall software
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key) as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                with winreg.OpenKey(key, subkey_name) as subkey:
                    try:
                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        programs.append(display_name)
                    except FileNotFoundError:
                        continue
    except Exception as e:
        print(f"Error accessing registry: {e}")
    return programs

def uninstall_program(program_name):
    """Uninstall the specified program using its uninstall string."""
    uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key) as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                with winreg.OpenKey(key, subkey_name) as subkey:
                    try:
                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        if display_name == program_name:
                            uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                            subprocess.run(uninstall_string, shell=True)
                            print(f"Uninstall command executed for: {program_name}")
                            return
                    except FileNotFoundError:
                        continue
    except Exception as e:
        print(f"Error accessing registry: {e}")

def recognize_speech():
    """Recognize speech input from the user."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Which program would you like to uninstall? Please say the number:")
        recognizer.adjust_for_ambient_noise(source) 
        audio = recognizer.listen(source)

    try:
        spoken_number = recognizer.recognize_google(audio)
        print(f"You said: {spoken_number}")
        return int(spoken_number)
    except sr.UnknownValueError:
        print("Sorry, I could not understand the audio.")
    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")
    except ValueError:
        print("Please say a valid number.")
    
    return None

def main():
    programs = list_installed_programs()
    if not programs:
        print("No programs found.")
        return

    print("Installed Programs:")
    for idx, program in enumerate(programs, start=1):
        print(f"{idx}. {program}")

    # Ask for verbal input
    choice = recognize_speech()
    if choice is not None and 1 <= choice <= len(programs):
        program_to_uninstall = programs[choice - 1]
        confirm = input(f"Confirm uninstallation of '{program_to_uninstall}'? (y/n): ")
        if confirm.lower() == 'y':
            uninstall_program(program_to_uninstall)
        else:
            print("Uninstallation canceled.")
    else:
        print("Invalid choice.")

if __name__ == "__main__":
    main()

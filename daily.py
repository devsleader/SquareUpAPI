import subprocess
import time

# Global variable to store the subprocess instance
current_process = None

def run_script(script_name):
    global current_process
    
    # If a previous process is running, terminate it
    if current_process is not None:
        print("Stopping previous execution...")
        current_process.terminate()
        current_process.wait()  # Wait for the process to terminate completely
        print("Previous execution stopped.")  
    
    # Start a new process for the given script
    print(f"Starting execution of {script_name}...")
    current_process = subprocess.Popen(["python", script_name])
    current_process.wait()  # Wait for the script to finish
    print(f"Execution of {script_name} completed.")
    
    # Add a 30-40 second delay after each script execution
    print(f"Waiting 10 Minuts before starting the next script...")
    time.sleep(20 * 60)  # You can adjust the time here to 10 Minuts if needed (e.g., time.sleep(40))

def run_scripts_in_sequence():
    scripts = [
        "1-cataLogFeedGoesHere.py",
        "1-openSheet.py",
        "2-Check_POS.py",
        "3-downloadSales.py"
    ]
    
    for script in scripts:
        run_script(script)

def start_scheduler(interval_minutes):
    while True:
        run_scripts_in_sequence()  # Run all scripts in sequence
        time.sleep(interval_minutes * 60)  # Wait for the specified interval

# Set the interval in minutes (e.g., 1440 minutes = 24 hours)
interval_minutes = 720 
print(f"Scheduler started. Running scripts every {interval_minutes} minutes.")
start_scheduler(interval_minutes)

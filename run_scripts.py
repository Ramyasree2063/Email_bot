import subprocess
import time

# Define the scripts to run in order
scripts = ["EmailBotMain.py"]

# Run the scripts continuously
while True:
    for script in scripts:
        subprocess.run(["python", script])
    time.sleep(10)  # Optional: wait for 10 seconds before the next run

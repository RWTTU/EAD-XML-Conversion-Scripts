# Navigate to the directory that contains your Python script
cd /path/to/your/script

# Install pyinstaller
pip install pyinstaller

# Create a standalone executable file using pyinstaller
pyinstaller --onefile your_script_name.py

# The executable file will be created in a dist directory in the same location as your script
# Distribute the executable file to users who don't have Python installed

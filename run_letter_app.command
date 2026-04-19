#!/bin/bash 
echo "Starting Letter Generator App..."
cd "$(dirname "$0")/letter_app"

# Setup virtual environment if missingk
if [ ! -d "venv" ]; then
    echo "Creating python environment..."
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt
else
    source venv/bin/activate
fi

# Run backend
echo "Opening in browser..."
sleep 2 && open "http://127.0.0.1:5050" &
python3 app.py

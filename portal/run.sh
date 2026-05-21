#!/bin/bash
set -e
cd "$(dirname "$0")"

echo "=== Care Weekly Dashboard Portal ==="

# Install dependencies if needed
if ! python3 -c "import fastapi" 2>/dev/null; then
  echo "Installing dependencies..."
  pip3 install -r requirements.txt -q
fi

echo "Starting portal at http://localhost:8000"
echo "Press Ctrl+C to stop."
python3 -m uvicorn app:app --host 0.0.0.0 --port 8000 --reload

import subprocess
import json
import sys

cmd = [r"C:\Users\Ramoncito\AppData\Local\Programs\Python\Python311\Scripts\notebooklm-mcp.exe"]
try:
    process = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
except Exception as e:
    print(f"Failed to start process: {e}")
    sys.exit(1)

def send_request(req):
    process.stdin.write(json.dumps(req) + "\n")
    process.stdin.flush()
    return process.stdout.readline()

# Initialize
init_req = {
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
        "protocolVersion": "2024-11-05",
        "capabilities": {},
        "clientInfo": {"name": "test", "version": "1.0"}
    }
}
print("Sending initialize...")
resp = send_request(init_req)
print("Init Response:", resp)

# Initialized
process.stdin.write(json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}) + "\n")
process.stdin.flush()

# List Resources
list_req = {
    "jsonrpc": "2.0",
    "id": 2,
    "method": "resources/list"
}
print("Listing resources...")
resp = send_request(list_req)
print("Resources:", resp)

process.terminate()

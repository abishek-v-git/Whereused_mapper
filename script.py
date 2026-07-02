import subprocess


result = subprocess.run(
    ["./deploy.sh"],
    capture_output=True,
    text=True
)

print(result.stdout)
print(result.stderr)
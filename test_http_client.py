import http.client
import urllib.parse
import json

boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
body = (
    f"--{boundary}\r\n"
    f'Content-Disposition: form-data; name="image"; filename="test.png"\r\n'
    f"Content-Type: image/png\r\n\r\n"
    f"fake_png_content\r\n"
    f"--{boundary}--\r\n"
)

conn = http.client.HTTPConnection("localhost", 8000)
headers = {
    "Content-Type": f"multipart/form-data; boundary={boundary}",
    "Content-Length": str(len(body))
}

conn.request("POST", "/api/gambarsoal", body.encode('utf-8'), headers)
response = conn.getresponse()
print("Status:", response.status)
print("Response:", response.read().decode())
conn.close()

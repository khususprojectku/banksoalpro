import urllib.request

data = b'''--boundary\r
Content-Disposition: form-data; name="image"; filename="test.png"\r
Content-Type: image/png\r
\r
fake_png_content\r
--boundary--\r
'''

req = urllib.request.Request(
    'http://localhost:8001/api/gambarsoal',
    data=data,
    headers={'Content-Type': 'multipart/form-data; boundary=boundary'},
    method='POST'
)

try:
    response = urllib.request.urlopen(req)
    print("STATUS:", response.status)
    print("BODY:", response.read().decode())
except urllib.error.HTTPError as e:
    print("STATUS:", e.code)
    print("BODY:", e.read().decode())

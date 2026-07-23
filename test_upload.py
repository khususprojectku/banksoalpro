import urllib.request
import urllib.error
import json

data = b'''--boundary\r
Content-Disposition: form-data; name="image"; filename="test.png"\r
Content-Type: image/png\r
\r
dummy_png_data\r
--boundary--\r
'''

req = urllib.request.Request(
    'http://localhost:8000/api/gambarsoal',
    data=data,
    headers={'Content-Type': 'multipart/form-data; boundary=boundary'}
)

try:
    response = urllib.request.urlopen(req)
    print(response.read().decode())
except urllib.error.HTTPError as e:
    print(e.code)
    print(e.read().decode())

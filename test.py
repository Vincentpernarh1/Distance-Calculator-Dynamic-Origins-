import requests

API_KEY = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6Ijk1OTNjMTAzMGViZDRkN2ZhMGNlMTk2ZGViMjFiOGIzIiwiaCI6Im11cm11cjY0In0="

url = "https://api.openrouteservice.org/v2/matrix/driving-car"

headers = {
    "Authorization": API_KEY,
    "Content-Type": "application/json; charset=utf-8",
    "Accept": "application/json, application/geo+json, application/gpx+xml, img/png; charset=utf-8",
    "User-Agent": "curl/8.0.1"
}

payload = {
    "locations": [
        [9.70093, 48.477473],
        [9.207916, 49.153868],
        [37.573242, 55.801281],
        [115.663757, 38.106467]
    ],
    "metrics": ["distance"]
}

r = requests.post(url, headers=headers, json=payload)

print("STATUS:", r.status_code)
print("RESPONSE:", r.text)

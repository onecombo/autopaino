import base64

text = "note_on"
encoded = base64.b64encode(text.encode('utf-8')).decode("ascii")
print(encoded)
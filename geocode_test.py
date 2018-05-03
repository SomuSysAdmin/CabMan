import googlemaps
import json

"""
Names of Related files
"""
api_key = "AIzaSyDP1Yy2gzWdrrB38GKPOEGiBj5B4I4sa1U"
gmaps = googlemaps.Client(api_key)

"""
res = gmaps.geocode("Biriyani Zone, Marathahalli, Bangalore")
print(json.dumps(res, indent=4))

Obtained:
=============
Addr: 92/5 ,Opposite Home Town Marathahalli, Marathahalli - Sarjapur Outer Ring Road, Bengaluru, Karnataka 560037, India
PlaceID:    ChIJw7OpJzUSrjsRNdQIDWszxHs
Co-Ords:    12.9523761, 77.7003376
"""

res2 = gmaps.reverse_geocode("ChIJw7OpJzUSrjsRNdQIDWszxHs")
print(json.dumps(res2, indent=4))
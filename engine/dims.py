import json, glob, os, sys
from PIL import Image
d = sys.argv[1].replace("\\", "/")
o = {}
for f in glob.glob(d + "/*.*"):
    try: o[os.path.basename(f)] = list(Image.open(f).size)
    except Exception: pass
print(json.dumps(o))
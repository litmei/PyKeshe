from PIL import Image
import os
import glob


files = glob.glob(r"E:\013_UniversitySubjectDesign\Data\R\00T_heater\*.png")
basewidth = 320
for f in files:
    title, ext = os.path.splitext(f)
    if not title.rfind("_half") == -1:
        continue
    img = Image.open(f)
    wpercent = basewidth/float(img.size[0])
    hsize = int((float(img.size[1]) * float(wpercent)))
    img_resize = img.resize((basewidth, hsize))
    img_resize.save(title + "_half" + ext, 'png')

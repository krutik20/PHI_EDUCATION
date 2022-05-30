import numpy as np
from PIL import Image

img = Image.open('bird.jpg')
red,green,blue = img.split()
redarr = np.asarray(red)
print(redarr ,"\n")
red_array_1d = redarr.flatten()
print(red_array_1d.ndim)
sorted_arr = np.sort(red_array_1d)
print(sorted_arr)
import matplotlib.pyplot as plt
​​​​​​​%matplotlib inline
 

from skimage import data,filters


image = data.coins()
# ... or any other NumPy array!
edges = filters.sobel(image)
plt.imshow(edges, cmap='gray')
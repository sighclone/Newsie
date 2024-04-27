import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import Normalize
from matplotlib.cm import ScalarMappable

# Read the Excel file
df = pd.read_excel('data_with_images.xlsx')

# Normalize subjectivity range to 0 to 1
df['subjectivity'] = (df['subjectivity'] - df['subjectivity'].min()) / (df['subjectivity'].max() - df['subjectivity'].min())

# Create a colormap based on 'subjectivity'
cmap = plt.get_cmap('coolwarm')

# Plot the data with colored dots
fig, ax = plt.subplots(figsize=(10, 6))
scatter = ax.scatter(df.index, df['polarity'], c=df['subjectivity'], cmap=cmap)
plt.colorbar(ScalarMappable(norm=Normalize(vmin=df['subjectivity'].min(), vmax=df['subjectivity'].max()), cmap=cmap), ax=ax, label='Subjectivity')
ax.set_xlabel('Index')
ax.set_ylabel('Polarity')
ax.set_title('Analysis Results')
plt.show()

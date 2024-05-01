import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

def pca(X):
 # mean center the data
 X_meaned = X - np.mean(X , axis = 0)
 
 # calculate the covariance matrix
 cov_mat = np.cov(X_meaned , rowvar = False)
 
 # calculate eigenvectors & eigenvalues of the covariance matrix
 eigen_values , eigen_vectors = np.linalg.eigh(cov_mat)
 
 # sort the eigenvalues in decreasing order
 sorted_index = np.argsort(eigen_values)[::-1]
 sorted_eigenvalue = eigen_values[sorted_index]
 sorted_eigenvectors = eigen_vectors[:,sorted_index]
 
 # select the first n eigenvectors, n is desired dimension
 # of our reduced subspace (i.e., number of principal components)
 eigenvector_subset = sorted_eigenvectors[:,0:4]

 # transform the data onto the new subspace
 X_reduced = np.dot(eigenvector_subset.transpose(),X_meaned.transpose() ).transpose()

 return X_reduced, sorted_eigenvalue, np.mean(X,axis=0)


excel_file = "C:\\Users\\rog\\OneDrive\\Desktop\\iiitb\\Sem4\\OutputExcel\\1\\DataFor1\\6.xlsx"
df = pd.read_excel(excel_file)
selected_columns = df.iloc[:, 1:9]
text_file = "output.txt"  # Name of the output text file
selected_columns.to_csv(text_file, sep=' ', index=False)
data = np.loadtxt(text_file, skiprows=1)
data2=[]
for i in range(len(data)):
 j=len(data[i])
 data1=[]
 for k in range(j):
  data1.append(data[i][k])
 data2.append(data1) #get all the data 
# Perform PCA
X_reduced, eigenvalues, mean = pca(data2)
# Print the variance explained by each principal component
variance_ratio = eigenvalues *100/ np.sum(eigenvalues)
print('Variance ratio: ', variance_ratio)
np.savetxt('afterPCA.txt', X_reduced)
data_after_pca = pd.read_csv('afterPCA.txt', header=None, delimiter=' ')

# Define the output Excel file name
excel_output_file = 'afterPCA.xlsx'

# Save the data to an Excel file
data_after_pca.to_excel(excel_output_file, index=False, header=False)
# Plot the data in the reduced subspace
plt.scatter(X_reduced[:, 0], X_reduced[:, 1])
plt.xlabel('PC1')
plt.ylabel('PC2')
plt.title('Plot after applying PCA')
plt.show()


import sys
import pandas as pd
import numpy as np
import toleranceinterval as ti
import scipy.stats as stats
from scipy.stats import norm
from numpy import sqrt
from scipy.stats import chi2

data = pd.read_excel(sys.argv[1], engine='openpyxl')

alpha = float(sys.argv[2]) #(1-a) since a=0.95

prop1 = float(sys.argv[3]) #this is the proportion for 1-sided-tol-int

prop2 = float(sys.argv[4]) #this is the proportion for the 2-sided-tol-int

data.columns = ["Unit Op", "Harvest" ,"Column1", "Viral Inactivation", "Column2","Column3","Viral Filtration","UFDF","DrugSubstance Fill",'Overall',      'Unnamed: 11',
            'Unnamed: 12',      'Unnamed: 13',      'Unnamed: 14',            'Unnamed: 15',      'Unnamed: 16',      'Unnamed: 17',
            'Unnamed: 18',      'Unnamed: 19',      'Unnamed: 20',
            'Unnamed: 21', "Unnamed: 22"]
data1 = data.tail(17)
data2 = data1.loc[~data1['Unit Op'].str.contains("Expt")]
realdata = data2[["Unit Op", "Harvest" ,"Column1", "Viral Inactivation", "Column2","Column3","Viral Filtration","UFDF","DrugSubstance Fill",'Overall']]

realdata1 = realdata.iloc[:, 1:]

results = pd.DataFrame(columns=['Column', 'Mean', 'Standard deviation', 'Lower tolerance limit', 'Upper tolerance limit'])
for col in realdata1:
	for i, j in realdata1[col].iteritems():
		if j < 1:
			realdata[col][i] = j * 100

# Calculate the mean, standard deviation, and confidence interval for each column
for col in realdata1.columns:


    col_mean = np.mean(realdata1[col])
    col_std = realdata1[col].std()
    n = realdata1[col].count()  # Count non-NaN values in the column
    # Below is the 2-sided tolerance interval
    dof = n - 1
    propinv = (1.0 - prop2) / 2.0
    z = norm.isf(propinv)
    chi = chi2.isf(q=(1-alpha), df=dof)
    k2 = sqrt((dof * (1 + (1/n)) * z**2)/ chi)
    lower_tol = col_mean - k2
    upper_tol = col_mean + k2
    # Below is the lower 1-sided tolerance interval
    zp = norm.isf((1-prop1))
    si =zp*sqrt(n)
    t = stats.t.ppf((1-alpha),dof,si)
    k1 = t / sqrt(n)      # Tolerance factor
    Tol_int = col_mean - (k1 * col_std)
    if Tol_int < 0:
    	Tol_int = Tol_int * -1
    print(f"Column: {col}")
    print(f"Mean: {col_mean}")
    print(f"Standard deviation: {col_std}")
    print(f"Tolerance interval (alpha={alpha}, prop1={prop1}, prop2={prop2}): ({lower_tol:.4f}, {upper_tol:.4f})")
    results = results.append({'Tol_int' : Tol_int ,'Column': col, 'Mean': col_mean, 'Standard deviation': col_std,'Count': n ,'Lower tolerance limit': lower_tol, 'Upper tolerance limit': upper_tol}, ignore_index=True)

results.to_excel(sys.argv[5], index=False)
print(f"All results returned to {sys.argv[5]}")
#output_file = sys.argv[4]
#columns = ['Column', 'Mean', 'Standard Deviation', 'Alpha', 'Proportion', 'Lower Tolerance Limit', 'Upper Tolerance Limit']
#results_df = pd.DataFrame(results, columns=columns)
#results_df.to_xlsx(output_file, index=False)



import sys
import pandas as pd
import numpy as np
import scipy.stats as stats
from scipy.stats import norm

data = pd.read_excel(sys.argv[1], engine='openpyxl')

alpha = float(sys.argv[2])

proportion = float(sys.argv[3])

data.columns = ["Unit Op", "Harvest" ,"Column1", "Viral Inactivation", "Column2","Column3","Viral Filtration","UFDF","DrugSubstance Fill",'Overall',      'Unnamed: 11',
            'Unnamed: 12',      'Unnamed: 13',      'Unnamed: 14',            'Unnamed: 15',      'Unnamed: 16',      'Unnamed: 17',
            'Unnamed: 18',      'Unnamed: 19',      'Unnamed: 20',
            'Unnamed: 21', "Unnamed: 22"]
data1 = data.tail(17)
data2 = data1.loc[~data1['Unit Op'].str.contains("Expt")]
realdata = data2[["Unit Op", "Harvest" ,"Column1", "Viral Inactivation", "Column2","Column3","Viral Filtration","UFDF","DrugSubstance Fill",'Overall']]

realdata1 = realdata.iloc[:, 1:]

results = pd.DataFrame(columns=['Column', 'Mean', 'Standard deviation', 'Lower tolerance limit', 'Upper tolerance limit'])


# Calculate the mean, standard deviation, and confidence interval for each column
for col in realdata1.columns:
    col_mean = np.mean(realdata1[col])
    col_std = np.std(realdata1[col])
    n = len(realdata1[col])
    z_alpha = norm.ppf(1 - alpha / 2)  # z-score for the given alpha level
    tol = z_alpha * col_std * np.sqrt(proportion / n)  # tolerance interval
    lower_tol = col_mean - tol  # lower tolerance limit
    upper_tol = col_mean + tol
    print(f"Column: {col}")
    print(f"Mean: {col_mean}")
    print(f"Standard deviation: {col_std}")
    print(f"Tolerance interval (alpha={alpha}, proportion={proportion}): ({lower_tol}, {upper_tol})\n")
    results = results.append({'Column': col, 'Mean': col_mean, 'Standard deviation': col_std, 'Lower tolerance limit': lower_tol, 'Upper tolerance limit': upper_tol}, ignore_index=True)

#output_file = sys.argv[4]
#columns = ['Column', 'Mean', 'Standard Deviation', 'Alpha', 'Proportion', 'Lower Tolerance Limit', 'Upper Tolerance Limit']
#results_df = pd.DataFrame(results, columns=columns)
#results_df.to_xlsx(output_file, index=False)

results.to_excel(sys.argv[4])

'''This script calculates the Gini Impurity for each feature in an Excel file.
 The smaller the result, the greater the impact of that feature.

Attention: The last column of the file must contain the values 'yes' or 'no'.'''

import pandas as pd

def calculate_gini(file_path):
    try:
        df = pd.read_excel(file_path)  # Read the Excel file
        label_col = df.columns[-1]  # Select the label column (last column)
    except: # Debugs file input problems and zero division problem except crashing
        print("It looks there is a problem, Please check the Excel file and try again.")
        return None

    gini_result = {}  # Save Gini results for each feature
    # Calculate Gini for each feature column
    for feature in df.columns[:-1]:  # Exclude the last column
        print("\033[1;32m"f"\nfor '{feature}'\033[0m") # Print in green("\033[1;32m")
        impurity_values = []  # List to Save Gini values for each unique value

        for value in df[feature].unique():  # "unique()" can ignore Duplicate values(from pandas)
            subset_labels = df[df[feature] == value][label_col]  # Labels for this value
            total = len(subset_labels)

            # Calculate the probability of 'yes'
            class_counts = subset_labels.value_counts()  # Counts 'yes' and 'no'
            prob_yes = 0 # Default / debug
            if 'yes' in class_counts:
                prob_yes = class_counts['yes'] / total
            prob_no = 1 - prob_yes # probability of 'no'

            # Calculate the Gini impurity using the probabilities
            impurity = 1 - (prob_yes ** 2 + prob_no ** 2)  # Gini mathematic formula
            impurity_values.append(impurity) # Save them as a list to use later
            # Print the probabilities and Gini of value in each feature (in yellow "\033[1;33m")
            print("\033[1;33m"f"    Value '{value}':""\033[0m" f" [Gini = {impurity:.4f}], [Prob = {prob_yes:.3f}]")
        gini_result[feature] = sum(impurity_values) / len(impurity_values) # Average Gini value for this feature
    print("\033[0;32m""-" * 55 + "\033[0m" )
    return gini_result

# Main order
file_path = "X:/Your/Excel/file/path/here.xlsx" # Path of the Excel file
gini_scores = calculate_gini(file_path)

# Prints final result for each feature
if gini_scores is not None: # Debugs empty columns issue
    print("\nFinal Gini Impurity results for features:")
    min_gini_scores = min(gini_scores.values())  # Find the smallest Gini score to make it red print
    for feature, gini in gini_scores.items():
        if gini == min_gini_scores:
            # Print the smallest Gini score in red("\033[1;31m")
            print("Feature""\033[1;31m" f" '{feature}'""\033[0m"" : Gini Impurity =""\033[1;31m" f" {gini:.4f}" + "\033[0m")
        else:
            # Print other Gini scores normally
            print(f"Feature '{feature}': Gini Impurity = {gini:.4f}")
input(">>>")



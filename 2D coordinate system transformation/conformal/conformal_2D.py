import pandas as pd
import numpy as np

# Read the Excel file
df = pd.read_excel('points_coordinate_system_1.xlsx')
df2 = pd.read_excel('control_points.xlsx')
df3 = pd.read_excel('old_coordinates.xlsx')
df4 = pd.read_excel('check1.xlsx')
df5 = pd.read_excel('check2.xlsx')

# Convert DataFrame to a list of lists (each row as a list)
points_coordinate_system_1 = df.values.tolist()
control_points = df2.values.tolist()
old_coordinates = df3.values.tolist()
chekc1 = df4.values.tolist()
chekc2 = df5.values.tolist()

# equations and unknowns
n = len(control_points) * 2
u = 4

# creating matrix of observations for least square algorithm
L = np.zeros((n, 1))

c = 0
for i in control_points:
    L[c][0] = i[0]
    c += 1
    L[c][0] = i[1]
    c += 1

# creating matrix A for least square algorithm
def create_A(arr, n, u):
    A = np.zeros((n, u))
    c = 0
    for i in arr:
        A[c][0] = i[0]
        A[c][1] = i[1]
        A[c][2] = 1
        A[c][3] = 0

        c += 1

        A[c][0] = i[1]
        A[c][1] = -i[0]
        A[c][2] = 0
        A[c][3] = 1

        c += 1
    return A

A = create_A(points_coordinate_system_1, n, u)

# calculating unknowns for 2D conformal transformation
X_CAP = np.linalg.inv(A.T @ A) @ A.T @ L

a = X_CAP[0][0]
b = X_CAP[1][0]
X0 = X_CAP[2][0]
Y0 = X_CAP[3][0]

# ----------------------------- RMSE CALCULATION -----------------------------
n = len(chekc1) * 2

A = create_A(chekc1, n, u)

new_check_coordinates = A @ X_CAP
new_check_coordinates = np.reshape(new_check_coordinates, newshape=(int(np.ceil(len(new_check_coordinates) / 2)), 2))

residuals = chekc2 - new_check_coordinates
residual_x = residuals[:, 0]
residual_y = residuals[:, 1]

sum_e = 0
for i in range(len(residual_x)):
    e = np.sqrt(residual_x[i] **2 + residual_y[i]**2)
    sum_e += (e**2)

rmse = np.sqrt(sum_e / len(residual_x))

print(f"Total RMSE: {rmse:.10f}")

# calculating scale
lambda_ = np.sqrt(a**2 + b**2)

# calculating kappa angle
K = np.rad2deg(np.arctan2(b, a))

# moving our coordinates to them new coordinates system
n = len(old_coordinates) * 2
A = create_A(old_coordinates, n, u)

new_coordinates = A @ X_CAP
new_coordinates = np.reshape(new_coordinates, newshape=(int(np.ceil(len(new_coordinates) / 2)), 2))

# creating csv file of our results
df4 = pd.DataFrame(new_coordinates, columns=["X_new", "Y_new"])
df4.to_csv(f"new_coordinates.csv", index=False)

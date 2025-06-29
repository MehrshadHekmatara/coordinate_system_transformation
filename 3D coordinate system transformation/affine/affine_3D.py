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
n = len(control_points) * 3
u = 12

# creating matrix of observations for least square algorithm
L = np.zeros((n, 1))

c = 0
for i in control_points:
    L[c][0] = i[0]
    c += 1
    L[c][0] = i[1]
    c += 1
    L[c][0] = i[2]
    c += 1

# creating matrix A for least square algorithm
def create_A(arr, n, u):
    A = np.zeros((n, u))
    c = 0
    for i in arr:
        A[c][0] = i[0]
        A[c][1] = i[1]
        A[c][2] = i[2]
        A[c][3] = 1
        A[c][4] = 0
        A[c][5] = 0
        A[c][6] = 0
        A[c][7] = 0
        A[c][8] = 0
        A[c][9] = 0
        A[c][10] = 0
        A[c][11] = 0

        c += 1

        A[c][0] = 0
        A[c][1] = 0
        A[c][2] = 0
        A[c][3] = 0
        A[c][4] = i[0]
        A[c][5] = i[1]
        A[c][6] = i[2]
        A[c][7] = 1
        A[c][8] = 0
        A[c][9] = 0
        A[c][10] = 0
        A[c][11] = 0

        c += 1

        A[c][0] = 0
        A[c][1] = 0
        A[c][2] = 0
        A[c][3] = 0
        A[c][4] = 0
        A[c][5] = 0
        A[c][6] = 0
        A[c][7] = 0
        A[c][8] = i[0]
        A[c][9] = i[1]
        A[c][10] = i[2]
        A[c][11] = 1

        c += 1
    return A

A = create_A(points_coordinate_system_1, n, u)

# calculating unknowns for 2D conformal transformation
X_CAP = np.linalg.inv(A.T @ A) @ A.T @ L

a1 = X_CAP[0][0]
a2 = X_CAP[1][0]
a3 = X_CAP[2][0]
a4 = X_CAP[3][0]
b1 = X_CAP[4][0]
b2 = X_CAP[5][0]
b3 = X_CAP[6][0]
b4 = X_CAP[7][0]
c1 = X_CAP[8][0]
c2 = X_CAP[9][0]
c3 = X_CAP[10][0]
c4 = X_CAP[11][0]

# ----------------------------- RMSE CALCULATION -----------------------------
n = len(chekc1) * 3

A = create_A(chekc1, n, u)

new_check_coordinates = A @ X_CAP
new_check_coordinates = np.reshape(new_check_coordinates, newshape=(int(np.ceil(len(new_check_coordinates) / 3)), 3))

residuals = chekc2 - new_check_coordinates
residual_x = residuals[:, 0]
residual_y = residuals[:, 1]
residual_z = residuals[:, 2]

sum_e = 0
for i in range(len(residual_x)):
    e = np.sqrt(residual_x[i] **2 + residual_y[i]**2 + residual_z[i]**2)
    sum_e += (e**2)

rmse = np.sqrt(sum_e / len(residual_x))

print(f"Total RMSE: {rmse:.10f}")

# moving our coordinates to them new coordinates system
n = len(old_coordinates) * 3

A = create_A(old_coordinates, n, u)

new_coordinates = A @ X_CAP
new_coordinates = np.reshape(new_coordinates, newshape=(int(np.ceil(len(new_coordinates) / 3)), 3))

# creating csv file of our results
df4 = pd.DataFrame(new_coordinates)
df4.to_csv(f"new_coordinates.csv", index=False)

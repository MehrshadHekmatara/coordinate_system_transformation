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
u = 15

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
def create_A(arr, cont, n, u):
    A = np.zeros((n, u))
    c = 0
    for i in range(len(arr)):
        A[c][0] = arr[i][0]
        A[c][1] = arr[i][1]
        A[c][2] = arr[i][2]
        A[c][3] = 1
        A[c][4] = 0
        A[c][5] = 0
        A[c][6] = 0
        A[c][7] = 0
        A[c][8] = 0
        A[c][9] = 0
        A[c][10] = 0
        A[c][11] = 0
        A[c][12] = -cont[i][0] * arr[i][0]
        A[c][13] = -cont[i][0] * arr[i][1]
        A[c][14] = -cont[i][0] * arr[i][2]

        c += 1

        A[c][0] = 0
        A[c][1] = 0
        A[c][2] = 0
        A[c][3] = 0
        A[c][4] = arr[i][0]
        A[c][5] = arr[i][1]
        A[c][6] = arr[i][2]
        A[c][7] = 1
        A[c][8] = 0
        A[c][9] = 0
        A[c][10] = 0
        A[c][11] = 0
        A[c][12] = -cont[i][1] * arr[i][0]
        A[c][13] = -cont[i][1] * arr[i][1]
        A[c][14] = -cont[i][1] * arr[i][2]

        c += 1

        A[c][0] = 0
        A[c][1] = 0
        A[c][2] = 0
        A[c][3] = 0
        A[c][4] = 0
        A[c][5] = 0
        A[c][6] = 0
        A[c][7] = 0
        A[c][8] = arr[i][0]
        A[c][9] = arr[i][1]
        A[c][10] = arr[i][2]
        A[c][11] = 1
        A[c][12] = -cont[i][2] * arr[i][0]
        A[c][13] = -cont[i][2] * arr[i][1]
        A[c][14] = -cont[i][2] * arr[i][2]

        c += 1
    return A

A = create_A(points_coordinate_system_1, control_points, n, u)

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
d1 = X_CAP[12][0]
d2 = X_CAP[13][0]
d3 = X_CAP[14][0]

# ----------------------------- RMSE CALCULATION -----------------------------
transformed_check = []

for point in chekc1:
    x, y, z = point[0], point[1], point[2]
    denominator = d1 * x + d2 * y + d3 * z + 1

    X_new = (a1 * x + a2 * y + a3 * z + a4) / denominator
    Y_new = (b1 * x + b2 * y + b3 * z + b4) / denominator
    Z_new = (c1 * x + c2 * y + c3 * z + c4) / denominator

    transformed_check.append([X_new, Y_new, Z_new])

residuals = np.array(chekc2) - np.array(transformed_check)
residual_x = residuals[:, 0]
residual_y = residuals[:, 1]
residual_z = residuals[:, 2]

sum_e = 0
for i in range(len(residual_x)):
    e = np.sqrt(residual_x[i] **2 + residual_y[i]**2 + residual_z[i]**2)
    sum_e += (e**2)

rmse = np.sqrt(sum_e / len(residual_x))

print(f"Total RMSE: {rmse:.10f}")

# Transform coordinates using the estimated projective transformation parameters
transformed_coords = []

for point in old_coordinates:
    x, y, z = point[0], point[1], point[2]
    denominator = d1 * x + d2 * y + d3 * z + 1

    X_new = (a1 * x + a2 * y + a3 * z + a4) / denominator
    Y_new = (b1 * x + b2 * y + b3 * z + b4) / denominator
    Z_new = (c1 * x + c2 * y + c3 * z + c4) / denominator

    transformed_coords.append([X_new, Y_new, Z_new])

# Convert the result to a DataFrame for display and saving
df_transformed = pd.DataFrame(transformed_coords, columns=['X_transformed', 'Y_transformed', 'z_transformed'])

# Save the transformed coordinates to an Excel file (optional)
df_transformed.to_excel('transformed_coordinates.xlsx', index=False)

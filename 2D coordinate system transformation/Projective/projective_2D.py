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
u = 8

# creating matrix of observations for least square algorithm
L = np.zeros((n, 1))

c = 0
for i in control_points:
    L[c][0] = i[0]
    c += 1
    L[c][0] = i[1]
    c += 1

# creating matrix A for least square algorithm
def create_A(arr, cont, n, u):
    A = np.zeros((n, u))
    c = 0
    for i in range(len(arr)):
        A[c][0] = arr[i][0]
        A[c][1] = arr[i][1]
        A[c][2] = 1
        A[c][3] = 0
        A[c][4] = 0
        A[c][5] = 0
        A[c][6] = -cont[i][0] * arr[i][0]
        A[c][7] = -cont[i][0] * arr[i][1]

        c += 1

        A[c][0] = 0
        A[c][1] = 0
        A[c][2] = 0
        A[c][3] = arr[i][0]
        A[c][4] = arr[i][1]
        A[c][5] = 1
        A[c][6] = -cont[i][1] * arr[i][0]
        A[c][7] = -cont[i][1] * arr[i][1]

        c += 1
    return A

A = create_A(points_coordinate_system_1, control_points, n, u)

# calculating unknowns for 2D conformal transformation
X_CAP = np.linalg.inv(A.T @ A) @ A.T @ L

a = X_CAP[0][0]
b = X_CAP[1][0]
c = X_CAP[2][0]
d = X_CAP[3][0]
e = X_CAP[4][0]
f = X_CAP[5][0]
g = X_CAP[6][0]
h = X_CAP[7][0]

# ----------------------------- RMSE CALCULATION -----------------------------
transformed_check = []

for point in chekc1:
    x, y = point[0], point[1]
    denominator = g * x + h * y + 1

    X_new = (a * x + b * y + c) / denominator
    Y_new = (d * x + e * y + f) / denominator

    transformed_check.append([X_new, Y_new])

residuals = np.array(chekc2) - np.array(transformed_check)
residual_x = residuals[:, 0]
residual_y = residuals[:, 1]

sum_e = 0
for i in range(len(residual_x)):
    e = np.sqrt(residual_x[i] **2 + residual_y[i]**2)
    sum_e += (e**2)

rmse = np.sqrt(sum_e / len(residual_x))

print(f"Total RMSE: {rmse:.10f}")

# Transform coordinates using the estimated projective transformation parameters
transformed_coords = []

for point in old_coordinates:
    x, y = point[0], point[1]
    denominator = g * x + h * y + 1

    X_new = (a * x + b * y + c) / denominator
    Y_new = (d * x + e * y + f) / denominator

    transformed_coords.append([X_new, Y_new])

# Convert the result to a DataFrame for display and saving
df_transformed = pd.DataFrame(transformed_coords, columns=['X_transformed', 'Y_transformed'])

# Save the transformed coordinates to an Excel file (optional)
df_transformed.to_excel('transformed_coordinates.xlsx', index=False)

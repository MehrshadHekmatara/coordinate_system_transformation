clc;clear;close all;

% Read Excel files
points_coordinate_system_1 = xlsread('points_coordinate_system_1.xlsx');
control_points = xlsread('control_points.xlsx');
old_coordinates = xlsread('old_coordinates.xlsx');
check1 = xlsread('check1.xlsx');
check2 = xlsread('check2.xlsx');

% Number of equations and unknowns
n = size(control_points, 1) * 3;
u = 12;

% Creating observation vector L for least squares
L = zeros(n, 1);
c = 1;
for i = 1:size(control_points, 1)
    L(c) = control_points(i, 1);
    c = c + 1;
    L(c) = control_points(i, 2);
    c = c + 1;
    L(c) = control_points(i, 3);
    c = c + 1;
end

% Create design matrix A using source points
A = create_A(points_coordinate_system_1, n, u);

% Least squares estimation of unknown parameters
X_CAP = inv(A' * A) * A' * L;

% Extracting transformation parameters
a1 = X_CAP(1);  a2 = X_CAP(2);  a3 = X_CAP(3);  a4 = X_CAP(4);
b1 = X_CAP(5);  b2 = X_CAP(6);  b3 = X_CAP(7);  b4 = X_CAP(8);
c1 = X_CAP(9);  c2 = X_CAP(10); c3 = X_CAP(11); c4 = X_CAP(12);

% ----------------------------- RMSE CALCULATION -----------------------------
% Apply transformation to source control points
n3 = size(check1, 1) * 3;
A3 = create_A(check1, n3, u);
new_check = A3 * X_CAP;
new_check = reshape(new_check, 3, [])';

% Calculate residuals between real and transformed control points
residuals = check2 - new_check;

residual_x = residuals(:, 1);
residual_y = residuals(:, 2);
residual_z = residuals(:, 3);

% Compute total RMSE
sum_e = 0;
for i = 1:length(residual_x)
    e = sqrt(residual_x(i)^2 + residual_y(i)^2 + residual_z(i)^2);
    sum_e = sum_e + e^2;
end
rmse = sqrt(sum_e / length(residual_x));
fprintf('Total RMSE: %.10f\n', rmse);

% ----------------------------- TRANSFORM OLD COORDINATES -----------------------------
% Rebuild matrix A for old coordinates
n = size(old_coordinates, 1) * 3;
A = create_A(old_coordinates, n, u);

% Apply transformation to old coordinates
new_coordinates = A * X_CAP;
new_coordinates = reshape(new_coordinates, 3, [])';

% Save new coordinates to CSV file
csvwrite('new_coordinates_matlab.csv', new_coordinates);

% Function to create design matrix A for 3D affine transformation
function A = create_A(arr, n, u)
    A = zeros(n, u);
    c = 1;
    for i = 1:size(arr, 1)
        x = arr(i, 1);
        y = arr(i, 2);
        z = arr(i, 3);

        A(c, :) = [x, y, z, 1, 0, 0, 0, 0, 0, 0, 0, 0];
        c = c + 1;

        A(c, :) = [0, 0, 0, 0, x, y, z, 1, 0, 0, 0, 0];
        c = c + 1;

        A(c, :) = [0, 0, 0, 0, 0, 0, 0, 0, x, y, z, 1];
        c = c + 1;
    end
end

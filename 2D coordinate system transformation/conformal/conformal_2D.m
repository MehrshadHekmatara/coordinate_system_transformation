clc;clear;close all;

% Read the Excel file
points_coordinate_system_1 = xlsread('points_coordinate_system_1.xlsx');
control_points = xlsread('control_points.xlsx');
old_coordinates = xlsread('old_coordinates.xlsx');
check1 = xlsread('check1.xlsx');
check2 = xlsread('check2.xlsx');

% Equations and unknowns
n = size(control_points, 1) * 2;
u = 4;

% Creating matrix of observations for least square algorithm
L = zeros(n, 1);

c = 1; % MATLAB indexing starts at 1
for i = 1:size(control_points, 1)
    L(c, 1) = control_points(i, 1);
    c = c + 1;
    L(c, 1) = control_points(i, 2);
    c = c + 1;
end

% Create matrix A
A = create_A(points_coordinate_system_1, n, u);

% Calculating unknowns for 2D conformal transformation
X_CAP = inv(A' * A) * A' * L;

% Compute new coordinates of control points using existing A and X_CAP
n3 = size(check1, 1) * 2;
A3 = create_A(check1, n3 ,u);
new_check = A3 * X_CAP;
new_check = reshape(new_check, 2, [])';

% Calculate residuals between real and transformed control points
residuals = check2 - new_check;

residual_x = residuals(:, 1);
residual_y = residuals(:, 2);

sum_e = 0;
for i = 1:length(residual_x)
    e = sqrt(residual_x(i)^2 + residual_y(i)^2);
    sum_e = sum_e + e^2;
end

rmse = sqrt(sum_e / length(residual_x));

fprintf('Total RMSE: %.10f\n', rmse);

a = X_CAP(1, 1);
b = X_CAP(2, 1);
X0 = X_CAP(3, 1);
Y0 = X_CAP(4, 1);

% Calculating scale
lambda = sqrt(a^2 + b^2);

% Calculating kappa angle
K = rad2deg(atan2(b, a));

% Moving our coordinates to the new coordinate system
n = size(old_coordinates, 1) * 2;
A = create_A(old_coordinates, n, u);
new_coordinates = A * X_CAP;
[s, ~] = size(new_coordinates);

final_coordinates = zeros(s/2, 2);

c = 1;
for i = 1:(s/2)
    final_coordinates(i, 1) = new_coordinates(c);
    c = c + 1;
    final_coordinates(i, 2) = new_coordinates(c);
    c = c + 1;
end

% Creating CSV file of our results
csvwrite('new_coordinates_matlab.csv', final_coordinates);

% Function to create matrix A for least square algorithm
function A = create_A(arr, n, u)
    A = zeros(n, u);
    c = 1; % MATLAB indexing starts at 1
    for i = 1:size(arr, 1)
        A(c, 1) = arr(i, 1);
        A(c, 2) = arr(i, 2);
        A(c, 3) = 1;
        A(c, 4) = 0;
        
        c = c + 1;
        
        A(c, 1) = arr(i, 2);
        A(c, 2) = -arr(i, 1);
        A(c, 3) = 0;
        A(c, 4) = 1;
        
        c = c + 1;
    end
end
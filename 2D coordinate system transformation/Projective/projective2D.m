clc;clear;close all;

% Read Excel files
df = xlsread('points_coordinate_system_1.xlsx');
df2 = xlsread('control_points.xlsx');
df3 = xlsread('old_coordinates.xlsx');
check1 = xlsread('check1.xlsx');
check2 = xlsread('check2.xlsx');

% Convert to lists (arrays)
points_coordinate_system_1 = df;
control_points = df2;
old_coordinates = df3;

% Equations and unknowns
n = size(control_points, 1) * 2;
u = 8;

% Creating matrix of observations for least square algorithm
L = zeros(n, 1);

c = 1;
for i = 1:size(control_points, 1)
    L(c) = control_points(i, 1);
    c = c + 1;
    L(c) = control_points(i, 2);
    c = c + 1;
end

% Creating matrix A for least square algorithm
A = zeros(n, u);
c = 1;
for i = 1:size(points_coordinate_system_1, 1)
    x = points_coordinate_system_1(i, 1);
    y = points_coordinate_system_1(i, 2);
    X = control_points(i, 1);
    Y = control_points(i, 2);

    A(c, :) = [x, y, 1, 0, 0, 0, -X * x, -X * y];
    c = c + 1;
    A(c, :) = [0, 0, 0, x, y, 1, -Y * x, -Y * y];
    c = c + 1;
end

% Solving for unknowns using least squares
X_CAP = inv(A' * A) * A' * L;

a = X_CAP(1);
b = X_CAP(2);
c_ = X_CAP(3); % renamed to c_ to avoid conflict with index variable
d = X_CAP(4);
e = X_CAP(5);
f = X_CAP(6);
g = X_CAP(7);
h = X_CAP(8);

% Compute new coordinates of control points using existing A and X_CAP
transformed_check = zeros(size(check1, 1), 2);

for i = 1:size(check1, 1)
    x = check1(i, 1);
    y = check1(i, 2);
    denominator = g * x + h * y + 1;

    X_new = (a * x + b * y + c_) / denominator;
    Y_new = (d * x + e * y + f) / denominator;

    transformed_check(i, :) = [X_new, Y_new];
end

% Calculate residuals between real and transformed control points
residuals = check2 - transformed_check;

residual_x = residuals(:, 1);
residual_y = residuals(:, 2);

sum_e = 0;
for i = 1:length(residual_x)
    e = sqrt(residual_x(i)^2 + residual_y(i)^2);
    sum_e = sum_e + e^2;
end

rmse = sqrt(sum_e / length(residual_x));

fprintf('Total RMSE: %.10f\n', rmse);

% Transform coordinates using the estimated projective transformation parameters
transformed_coords = zeros(size(old_coordinates, 1), 2);

for i = 1:size(old_coordinates, 1)
    x = old_coordinates(i, 1);
    y = old_coordinates(i, 2);
    denominator = g * x + h * y + 1;

    X_new = (a * x + b * y + c_) / denominator;
    Y_new = (d * x + e * y + f) / denominator;

    transformed_coords(i, :) = [X_new, Y_new];
end

% Save the transformed coordinates to an Excel file
xlswrite('transformed_coordinates_matlab.xlsx', transformed_coords);

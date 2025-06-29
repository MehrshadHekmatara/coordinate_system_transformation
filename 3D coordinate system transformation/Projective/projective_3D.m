clc;clear;close all;

% Read Excel files
points_coordinate_system_1 = xlsread('points_coordinate_system_1.xlsx');
control_points = xlsread('control_points.xlsx');
old_coordinates = xlsread('old_coordinates.xlsx');
check1 = xlsread('check1.xlsx');
check2 = xlsread('check2.xlsx');

% equations and unknowns
n = size(control_points, 1) * 3;
u = 15;

% Create observation matrix L
L = zeros(n, 1);
c = 1;
for i = 1:size(control_points, 1)
    L(c) = control_points(i, 1); c = c + 1;
    L(c) = control_points(i, 2); c = c + 1;
    L(c) = control_points(i, 3); c = c + 1;
end

A = create_A(points_coordinate_system_1, control_points, n, u);

% Solve for unknowns using Least Squares
X_CAP = inv(A' * A) * A' * L;

a1 = X_CAP(1);  a2 = X_CAP(2);  a3 = X_CAP(3);  a4 = X_CAP(4);
b1 = X_CAP(5);  b2 = X_CAP(6);  b3 = X_CAP(7);  b4 = X_CAP(8);
c1 = X_CAP(9);  c2 = X_CAP(10); c3 = X_CAP(11); c4 = X_CAP(12);
d1 = X_CAP(13); d2 = X_CAP(14); d3 = X_CAP(15);

% ---------------------- RMSE CALCULATION ----------------------
transformed_check = [];

for i = 1:size(check1,1)
    x = check1(i,1);
    y = check1(i,2);
    z = check1(i,3);

    denominator = d1 * x + d2 * y + d3 * z + 1;

    X_new = (a1 * x + a2 * y + a3 * z + a4) / denominator;
    Y_new = (b1 * x + b2 * y + b3 * z + b4) / denominator;
    Z_new = (c1 * x + c2 * y + c3 * z + c4) / denominator;

    transformed_check = [transformed_check; X_new, Y_new, Z_new];
end

residuals = check2 - transformed_check;

residual_x = residuals(:,1);
residual_y = residuals(:,2);
residual_z = residuals(:,3);

sum_e = 0;
for i = 1:length(residual_x)
    e = sqrt(residual_x(i)^2 + residual_y(i)^2 + residual_z(i)^2);
    sum_e = sum_e + e^2;
end

rmse = sqrt(sum_e / length(residual_x));
fprintf('Total RMSE: %.10f\n', rmse);

% ---------------------- Transform Old Coordinates ----------------------
transformed_coords = [];

for i = 1:size(old_coordinates,1)
    x = old_coordinates(i,1);
    y = old_coordinates(i,2);
    z = old_coordinates(i,3);

    denominator = d1 * x + d2 * y + d3 * z + 1;

    X_new = (a1 * x + a2 * y + a3 * z + a4) / denominator;
    Y_new = (b1 * x + b2 * y + b3 * z + b4) / denominator;
    Z_new = (c1 * x + c2 * y + c3 * z + c4) / denominator;

    transformed_coords = [transformed_coords; X_new, Y_new, Z_new];
end

% Save result to Excel
xlswrite('transformed_coordinates_matlab.xlsx', transformed_coords);

% Create design matrix A
function A = create_A(arr, cont, n, u)
    A = zeros(n, u);
    c = 1;
    for i = 1:size(arr, 1)
        A(c, 1) = arr(i,1);
        A(c, 2) = arr(i,2);
        A(c, 3) = arr(i,3);
        A(c, 4) = 1;
        A(c, 13) = -cont(i,1) * arr(i,1);
        A(c, 14) = -cont(i,1) * arr(i,2);
        A(c, 15) = -cont(i,1) * arr(i,3);
        c = c + 1;

        A(c, 5) = arr(i,1);
        A(c, 6) = arr(i,2);
        A(c, 7) = arr(i,3);
        A(c, 8) = 1;
        A(c, 13) = -cont(i,2) * arr(i,1);
        A(c, 14) = -cont(i,2) * arr(i,2);
        A(c, 15) = -cont(i,2) * arr(i,3);
        c = c + 1;

        A(c, 9) = arr(i,1);
        A(c,10) = arr(i,2);
        A(c,11) = arr(i,3);
        A(c,12) = 1;
        A(c, 13) = -cont(i,3) * arr(i,1);
        A(c, 14) = -cont(i,3) * arr(i,2);
        A(c, 15) = -cont(i,3) * arr(i,3);
        c = c + 1;
    end
end

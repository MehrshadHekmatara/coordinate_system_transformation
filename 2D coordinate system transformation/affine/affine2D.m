clc;clear;close all;

% Read Excel files
points_coordinate_system_1 = xlsread('points_coordinate_system_1.xlsx');
control_points = xlsread('control_points.xlsx');
old_coordinates = xlsread('old_coordinates.xlsx');
check1 = xlsread('check1.xlsx');
check2 = xlsread('check2.xlsx');

% Number of equations and unknowns
n = size(control_points, 1) * 2;
u = 6;

% Create observation vector L
L = zeros(n, 1);
c = 1;
for i = 1:size(control_points, 1)
    L(c, 1) = control_points(i, 1);
    c = c + 1;
    L(c, 1) = control_points(i, 2);
    c = c + 1;
end

% Construct matrix A for control points
A = create_A(points_coordinate_system_1);

% Solve least squares equation
X_CAP = inv(A' * A) * A' * L;

% Extract transformation parameters
a = X_CAP(1);
b = X_CAP(2);
c = X_CAP(3);
d = X_CAP(4);
e = X_CAP(5);
f = X_CAP(6);

% Compute new coordinates of control points using existing A and X_CAP
n3 = size(check1, 1) * 2;
A3 = create_A(check1);
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

% Compute scales
lambda_x = sqrt(a^2 + c^2);
lambda_y = sqrt(b^2 + d^2);

% Compute rotation angles
B = rad2deg(atan2(-c, a));
alpha = rad2deg(atan2((a*b + c*d), (a*d - c*b)));

% Translations
xt = e;
yt = f;

% Apply transformation to old coordinates
n2 = size(old_coordinates, 1) * 2;
A2 = create_A(old_coordinates);
new_coordinates = A2 * X_CAP;
new_coordinates = reshape(new_coordinates, 2, [])';

% Save result to CSV
csvwrite('new_coordinates_matlab.csv', new_coordinates);

% Initialize matrix A
function A = create_A(arr)
    % arr: Nx2 matrix of points
    % Returns A: (2N)x6 matrix for 2D affine least squares

    n = size(arr, 1) * 2;  % number of rows
    u = 6;                 % number of unknowns
    A = zeros(n, u);
    
    c = 1;
    for i = 1:size(arr,1)
        x = arr(i,1);
        y = arr(i,2);
        
        % First row for point i
        A(c, 1) = x;
        A(c, 2) = y;
        A(c, 3) = 0;
        A(c, 4) = 0;
        A(c, 5) = 1;
        A(c, 6) = 0;
        c = c + 1;
        
        % Second row for point i
        A(c, 1) = 0;
        A(c, 2) = 0;
        A(c, 3) = x;
        A(c, 4) = y;
        A(c, 5) = 0;
        A(c, 6) = 1;
        c = c + 1;
    end
end


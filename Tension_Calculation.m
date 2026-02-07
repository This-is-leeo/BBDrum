

clear; clc;

%% settings
xlsxFile    = "parameter.xlsx";
inputSheet  = "input";
outputSheet = "Results";
materialSheet = "mats";

%% read + map rows into a struct + Material
raw = readcell(xlsxFile,"Sheet",inputSheet);
hdr = find(strcmpi(string(raw(:,1)),"Variable Name"),1,"first");
data = raw(hdr+1:end, 1:3); % [Variable Name, Value, Unit]

mats_raw = readcell(xlsxFile,"Sheet",materialSheet);
mats_headers = mats_raw(1, :);
mats_data    = mats_raw(2:end, :);
vars = struct();
for i = 1:size(data,1)
    name = data{i,1};
    % skip empty or missing rows
    if isempty(name) || (isstring(name) && ismissing(name))
        continue
    end

    if ismissing(name)
      key = "empty"; % or a default string
    else
      key = matlab.lang.makeValidName(lower(string(name)));
    end
    vars.(key) = struct("value",data{i,2},"unit",string(data{i,3}));
end


materials = struct([]);

for i = 1:size(mats_data,1)
    materials(i).material         = mats_data{i,1};
    materials(i).density          = double(mats_data{i,2});
    materials(i).tensileStrength  = double(mats_data{i,3});
    materials(i).youngsModulus    = double(mats_data{i,4});
end %adds material

%% pull inputs
f = double(vars.frequency.value);              % Hz
D = double(vars.diameter.value);               % in units below
Dunit = lower(strtrim(vars.diameter.unit));    % "ft" or "m"
if Dunit == "ft", D = D*0.3048; end            % -> meters
a = D/2;                                       % radius (m)
%material_type = 
alpha = double(vars.amn.value);   %Constant   

n = vars.fos.value;                             %Factor of Safety
h_mm = double(vars.membrane_thickness.value);  % mm
h = h_mm*1e-3;                                 % m
                                   % kg/m^2
%Materials
ID = double(vars.membrane_type.value);
rho = materials(ID).density;     %rho
sigma_u_MPa = materials(ID).tensileStrength;    %Tensile Strength

%calculations
mu = rho*h; 
sigma_u = sigma_u_MPa * 1e6;   % [Pa]
T_max   = sigma_u * h;         % [N/m]
T_allow = T_max / n;           % [N/m]
fprintf('T_max   = %.3g N/m (%.3g kN/m)\n', T_max,   T_max/1e3);
fprintf('T_allow = %.3g N/m (%.3g kN/m)  (SF=%g)\n', T_allow, T_allow/1e3, n);

%% 3D sweep: T vs pitch (Hz) and diameter (m)
% Uses: alpha, mu already computed in script

pitch = linspace(10,110,30);      % Hz 
Dvec  = linspace(5.5,6.5,30);     % m  

[P, D] = meshgrid(pitch, Dvec);
a = D/2;

Tsurf = mu .* ((2*pi.*a.*P)./alpha).^2;   % N/m

figure;
surf(P, D, Tsurf); hold on;

surf(P, D, T_allow*ones(size(P)), ...
    'FaceAlpha',0.55, ...
    'EdgeColor','none', 'faceColor', 'red'); hold on;

contour3(P, D, Tsurf, [T_allow T_allow], ...
    'k', 'LineWidth', 3);

xlabel("Pitch (Hz)");
ylabel("Diameter (m)");
zlabel("Tension T (N/m)");
title("Drumhead Tension vs Pitch and Diameter of " + string(materials(ID).material));
grid on;


%% calculate
T    = mu * ((2*pi*a*f)/alpha)^2/1000;  % kN/m
Frim = 2*pi*a*T;                   % kN

%% write results
out = {
    "Tension_T", T, "kN/m";
    "RimForce",  Frim, "kN";
};
disp(out)
%writecell(out, xlsxFile, "Sheet", outputSheet, "Range", "A1");

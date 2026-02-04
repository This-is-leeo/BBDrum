%% calc_drum_tension_from_excel.m  (SIMPLE ONE-FILE SCRIPT)
% Excel columns: Variable Name | Value | Unit | Comments
% Inputs expected (Variable Name):
%   frequency (Hz)
%   diameter (ft or m; set Unit cell)
%   membrane_density (PVC or numeric)   % PVC -> uses rho = 1400 kg/m^3
%   membrane_thickness (mm)            % used with rho to compute mu
%   Amn (optional)                     % default 2.4048
%
% Output written to new sheet "Results": T (N/m), Frim (N)

clear; clc;

%% settings
xlsxFile    = "parameter.xlsx";
inputSheet  = "input";
outputSheet = "Results";

%% read + map rows into a struct
raw = readcell(xlsxFile,"Sheet",inputSheet);
hdr = find(strcmpi(string(raw(:,1)),"Variable Name"),1,"first");
data = raw(hdr+1:end, 1:3); % [Variable Name, Value, Unit]

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

%% pull inputs
f = double(vars.frequency.value);              % Hz
D = double(vars.diameter.value);               % in units below
Dunit = lower(strtrim(vars.diameter.unit));    % "ft" or "m"
if Dunit == "ft", D = D*0.3048; end            % -> meters
a = D/2;                                       % radius (m)

alpha = double(vars.amn.value);                %Constant   
rho = double(vars.membrane_density.value);     %rho
% mu (kg/m^2)
h_mm = double(vars.membrane_thickness.value);  % mm
h = h_mm*1e-3;                                 % m
mu = rho*h;                                    % kg/m^2

%% calculate
T    = mu * ((2*pi*a*f)/alpha)^2/1000;  % kN/m
Frim = 2*pi*a*T;                   % kN

%% write results
out = {
    "Tension_T", T, "kN/m";
    "RimForce",  Frim, "kN";
};
disp(out)
writecell(out, xlsxFile, "Sheet", outputSheet, "Range", "A1");

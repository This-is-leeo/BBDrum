%% Tension_Calculation.m (SCRIPT)
% Reads an Excel file formatted like your screenshot:
% Columns: Variable Name | Value | Unit | Comments
% Computes membrane tension (N/m) from frequency and writes results to a new sheet.

clear; clc;

%% ---- USER SETTINGS ----
xlsxFile    = "parameter.xlsx";  % <--- change this
inputSheet  = "input";                   % can be sheet index or "Sheet1"
outputSheet = "Results";           % new sheet name
%% -----------------------

% Read raw cells so we can handle mixed text/numbers easily
raw = readcell(xlsxFile, "Sheet", inputSheet);

% Find header row containing "Variable Name"
headerRow = findRowContaining(raw, "Variable Name");
if isempty(headerRow)
    error('Could not find a header row containing "Variable Name".');
end

% Get column indices based on header names
headers = raw(headerRow, :);
colVar  = findCol(headers, "Variable Name");
colVal  = findCol(headers, "Value");
colUnit = findCol(headers, "Unit");

if any([colVar colVal colUnit] == 0)
    error('Expected columns: Variable Name, Value, Unit (Comments optional).');
end

% Parse key-value pairs below header
vars = struct();
for r = headerRow+1:size(raw,1)
    vname = raw{r, colVar};
    if ismissingCell(vname), continue; end

    vname = string(vname);
    if strlength(strtrim(vname))==0
        continue;
    end

    val  = raw{r, colVal};
    unit = "";
    if colUnit > 0 && colUnit <= size(raw,2) && ~ismissingCell(raw{r,colUnit})
        unit = string(raw{r, colUnit});
    end

    key = matlab.lang.makeValidName(lower(strtrim(vname)));
    vars.(key) = struct("value", val, "unit", unit);
end

%% ---- Pull required inputs ----
f = getNumeric(vars, "frequency"); % Hz
if isnan(f), error('Missing or non-numeric "frequency".'); end

diameter_m = getLengthInMeters(vars, "diameter"); % supports ft/in/m/cm/mm
if isnan(diameter_m), error('Missing or non-numeric "diameter".'); end
a = diameter_m/2; % radius (m)

% Alpha (Bessel root): use Amn if present, else fundamental 2.4048
alpha = getNumeric(vars, "amn");
if isnan(alpha), alpha = 2.4048; end

% Areal mass density mu (kg/m^2) — needed to compute tension
mu = computeArealDensity(vars);
if isnan(mu) || mu <= 0
    error(['Could not compute membrane areal density μ (kg/m^2). ' ...
           'Provide membrane_thickness (mm) and membrane_density (PVC or rho), ' ...
           'OR provide membrane_density directly as numeric kg/m^2 with unit kg/m^2.']);
end

%% ---- Physics ----
% f = (alpha/(2*pi*a)) * sqrt(T/mu)  =>  T = mu * (2*pi*a*f/alpha)^2
T    = mu * ((2*pi*a*f)/alpha)^2; % N/m  (tension per unit length)
c    = sqrt(T/mu);                % m/s  membrane wave speed
Frim = 2*pi*a*T;                  % N    total rim load (inward hoop force)

%% ---- Write results ----
out = {
    "Variable","Value","Unit","Notes";
    "frequency", f, "Hz", "";
    "diameter", diameter_m, "m", "Converted to meters";
    "radius", a, "m", "";
    "alpha_mn", alpha, "-", "Bessel root used";
    "mu_areal_density", mu, "kg/m^2", "Membrane areal mass density";
    "wave_speed_c", c, "m/s", "c = sqrt(T/mu)";
    "tension_T", T, "N/m", "Surface tension (per unit length)";
    "rim_force_total", Frim, "N", "Frim = 2*pi*a*T";
};

meta = {
    "GeneratedOn", string(datetime('now')), "", ""
};

writecell(meta, xlsxFile, "Sheet", outputSheet, "Range", "A1");
writecell(out,  xlsxFile, "Sheet", outputSheet, "Range", "A3");

fprintf('Wrote results to sheet "%s" in %s\n', outputSheet, xlsxFile);

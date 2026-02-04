%% drum_excel_helpers.m
% Helper functions used by calc_drum_tension_from_excel.m
% Put this file in the same folder (or on MATLAB path).

function r = findRowContaining(raw, needle)
    r = [];
    needle = lower(string(needle));
    for i=1:size(raw,1)
        row = raw(i,:);
        for j=1:numel(row)
            if ~ismissingCell(row{j})
                s = lower(string(row{j}));
                if contains(s, needle)
                    r = i; return;
                end
            end
        end
    end
end

function c = findCol(headers, name)
    c = 0;
    name = lower(string(name));
    for j=1:numel(headers)
        if ~ismissingCell(headers{j})
            if lower(string(headers{j})) == name
                c = j; return;
            end
        end
    end
end

function tf = ismissingCell(x)
    tf = isempty(x) || (isstring(x) && strlength(x)==0) || (ischar(x) && isempty(strtrim(x))) || ...
         (isnumeric(x) && isnan(x));
end

function x = getNumeric(vars, key)
    key = matlab.lang.makeValidName(lower(key));
    x = NaN;
    if ~isfield(vars, key), return; end
    v = vars.(key).value;

    if isnumeric(v)
        x = double(v);
        return;
    end
    if isstring(v) || ischar(v)
        s = strtrim(string(v));
        tok = regexp(s, '[-+]?\d*\.?\d+([eE][-+]?\d+)?', 'match', 'once');
        if ~isempty(tok), x = str2double(tok); end
    end
end

function Lm = getLengthInMeters(vars, key)
    key = matlab.lang.makeValidName(lower(key));
    Lm = NaN;
    if ~isfield(vars, key), return; end

    val  = vars.(key).value;
    unit = "";
    if isfield(vars.(key), "unit"), unit = lower(strtrim(string(vars.(key).unit))); end

    % If user typed "20 ft" in the Value cell, parse that too
    if (isstring(val) || ischar(val)) && strlength(strtrim(string(val)))>0
        s = lower(strtrim(string(val)));
        num = regexp(s, '[-+]?\d*\.?\d+([eE][-+]?\d+)?', 'match', 'once');
        if ~isempty(num)
            valNum = str2double(num);
            % attempt unit from the string
            if contains(s,"ft") || contains(s,"feet"), unit = "ft"; end
            if contains(s,"in"), unit = "in"; end
            if contains(s,"mm"), unit = "mm"; end
            if contains(s,"cm"), unit = "cm"; end
            if contains(s,"m"),  unit = "m";  end
            val = valNum;
        end
    end

    if ~isnumeric(val) || isnan(val)
        return;
    end

    val = double(val);
    if unit == "ft"
        Lm = val * 0.3048;
    elseif unit == "in"
        Lm = val * 0.0254;
    elseif unit == "mm"
        Lm = val * 1e-3;
    elseif unit == "cm"
        Lm = val * 1e-2;
    elseif unit == "m" || unit == ""
        Lm = val; % assume meters if unit blank
    else
        % Unknown unit: assume meters
        Lm = val;
    end
end

function mu = computeArealDensity(vars)
    % Returns mu in kg/m^2.
    % Supports:
    % - membrane_density numeric with unit kg/m^2 -> mu directly
    % - membrane_density = "PVC" (or numeric rho) + membrane_thickness -> mu = rho*h

    mu = NaN;

    rho = NaN;

    if isfield(vars, "membrane_density")
        md = vars.membrane_density.value;
        u  = lower(strtrim(string(vars.membrane_density.unit)));

        if isnumeric(md) && ~isnan(md)
            if contains(u,"kg/m^2") || contains(u,"kg/m2")
                mu = double(md);
                return;
            else
                rho = double(md); % assume kg/m^3 if not specified
            end
        else
            mat = lower(strtrim(string(md)));
            if mat == "pvc"
                rho = 1400; % kg/m^3 typical PVC
            elseif mat == "pet" || mat == "mylar"
                rho = 1390;
            elseif mat == "latex"
                rho = 920;
            end
        end
    end

    h = NaN;
    if isfield(vars, "membrane_thickness")
        h = getLengthInMeters(vars, "membrane_thickness"); % supports mm/cm/m
    end

    if ~isnan(rho) && ~isnan(h) && h > 0
        mu = rho * h;
    end
end

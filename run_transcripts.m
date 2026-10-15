%% run_transcripts.m
% Beyond Academics transcript automation (OSB, AUB)
%
% For every new Microsoft Forms response:
%   1. checks the student is OSB and agreed to the terms (otherwise logs it
%      to review_log.csv and skips it)
%   2. opens that student's transcript in OneDrive, or creates it from the
%      official template on their first submission
%   3. fills the header (name, concentration, date issued) and adds the
%      activity to the matching category table
%   4. saves the file back to OneDrive
%
% Already-processed responses are remembered in processed_log.csv, so running
% this again never adds the same activity twice.
%
% All settings are in ba_config.m.

clear; clc;
cfg = ba_config();
fprintf('Beyond Academics transcript automation\n');
if cfg.dryRun, fprintf('*** DRY RUN: nothing will be uploaded ***\n'); end

%% 0. Sign in to Microsoft 365 (first run: follow the printed instructions)
g = GraphClient(cfg.tenantId, cfg.clientId, cfg.scopes, cfg.tokenCache);
g.login();

work = fullfile(tempdir, 'ba_work');           % local scratch space
if isfolder(work), rmdir(work, 's'); end
mkdir(work);
cleanupWork = onCleanup(@() ba_safe_rmdir(work));

%% 1. Read the Forms responses workbook from OneDrive
fprintf('Reading responses: %s\n', cfg.responsesPath);
xbytes = g.download(cfg.responsesPath);
if isempty(xbytes)
    error('Responses workbook not found in OneDrive at "%s". Check ba_config.m.', cfg.responsesPath);
end
xlsxPath = fullfile(work, 'responses.xlsx');
ba_write_bytes(xlsxPath, xbytes);
T = readtable(xlsxPath, 'VariableNamingRule', 'preserve');
fprintf('Found %d response rows.\n', height(T));

% Forms uses the full question text as the column header, so match on a
% distinctive part of each question.
col.id            = ba_find_column(T, {'ID'}, true);
col.aubId         = ba_find_column(T, {'AUB ID'});
col.fullName      = ba_find_column(T, {'Full Name'});
col.semester      = ba_find_column(T, {'semester you have participated', 'specify in which semester'});
col.isOSB         = ba_find_column(T, {'Are you an OSB student'});
col.concentration = ba_find_column(T, {'Choose your concentration'});
col.activity      = ba_find_column(T, {'full activity name'});
col.category      = ba_find_column(T, {'category that best describes'});
col.duration      = ba_find_column(T, {'duration of your extracurricular'});
col.distinction   = ba_find_column(T, {'distinction or notable achievement'});
col.terms         = ba_find_column(T, {'agree to the terms and conditions'});

required = setdiff(fieldnames(col), {'id'});
missing = required(cellfun(@(f) isempty(col.(f)), required));
if ~isempty(missing)
    fprintf('Columns in the workbook:\n'); disp(T.Properties.VariableNames');
    error('Could not find these questions in the workbook: %s', strjoin(missing, ', '));
end

%% 2. Load logs and the template
processedPath = [cfg.stateFolder '/processed_log.csv'];
reviewPath    = [cfg.stateFolder '/review_log.csv'];
processed = ba_read_lines(g.download(processedPath));
reviewTxt = ba_utf8(g.download(reviewPath));
if isempty(reviewTxt)
    reviewTxt = sprintf('run_time,response,student,reason\n');
end

tplBytes = g.download(cfg.templatePath);
if isempty(tplBytes)
    error('Transcript template not found in OneDrive at "%s". Check ba_config.m.', cfg.templatePath);
end

%% 3. Process new responses
runTime = char(datetime('now', 'Format', 'yyyy-MM-dd HH:mm'));
dateIssued = char(datetime('now', 'Format', cfg.dateFormat));
nDone = 0; nSkipped = 0; nFailed = 0; nOld = 0;

for i = 1:height(T)
    v = @(c) ba_cell_text(T, c, i);
    name  = ba_clean_name(v(col.fullName));
    aubId = regexprep(v(col.aubId), '\s', '');
    if isempty(col.id)
        key = [v(col.aubId) '|' v(col.activity) '|' v(col.semester)];
    else
        key = ['ID ' v(col.id)];
    end
    if isempty(name) && isempty(aubId), continue; end           % empty row
    if any(strcmp(processed, key)), nOld = nOld + 1; continue; end
    who = sprintf('%s (%s)', name, aubId);

    % --- checks -----------------------------------------------------------
    reason = '';
    if ~startsWith(lower(strtrim(v(col.isOSB))), 'yes')
        reason = 'Not an OSB student';
    elseif ~ba_is_agreement(v(col.terms))
        reason = 'Did not agree to terms';
    elseif isempty(regexp(aubId, '^[0-9]{6,}$', 'once'))
        reason = 'AUB ID missing or not a number';
    end
    heading = ba_category_heading(v(col.category));
    if isempty(reason) && isempty(heading)
        reason = ['Unknown category: ' v(col.category)];
    end
    if ~isempty(reason)
        fprintf('  SKIP  %s: %s\n', who, reason);
        reviewTxt = [reviewTxt ba_csv_line({runTime, key, who, reason})]; %#ok<AGROW>
        processed{end+1} = key; %#ok<SAGROW>
        nSkipped = nSkipped + 1;
        continue;
    end

    % --- update the student's file ---------------------------------------
    title = strtrim(v(col.activity));
    dist = strtrim(v(col.distinction));
    if ~isempty(dist) && isempty(regexpi(dist, '^(none|no|n/?a|-+|\.)$', 'once'))
        title = [title ' - ' dist];
    end
    fileName = [regexprep(name, '\s+', '_') '_' aubId '.docx'];
    drivePath = [cfg.outputFolder '/' fileName];

    try
        bytes = g.download(drivePath);
        isNew = isempty(bytes);
        if isNew, bytes = tplBytes; end
        docx = fullfile(work, 'student.docx');
        ba_write_bytes(docx, bytes);
        unz = fullfile(work, 'unzipped');
        xml = TranscriptDocx.readXml(docx, unz);
        xml = TranscriptDocx.setHeader(xml, name, strtrim(v(col.concentration)), dateIssued);
        [xml, nEntries] = TranscriptDocx.addEntry(xml, heading, strtrim(v(col.semester)), ...
                                                  title, strtrim(v(col.duration)));
        TranscriptDocx.writeDocx(unz, xml, docx);

        if ~cfg.dryRun
            g.upload(drivePath, ba_read_bytes(docx));
        end
        if nEntries > cfg.visibleRowsPerCategory
            msg = sprintf('%s has %d entries (box shows %d) - check layout', ...
                          heading, nEntries, cfg.visibleRowsPerCategory);
            reviewTxt = [reviewTxt ba_csv_line({runTime, key, who, msg})]; %#ok<AGROW>
        end
        processed{end+1} = key; %#ok<SAGROW>
        nDone = nDone + 1;
        if isNew, tag = 'NEW '; else, tag = 'ADD '; end
        fprintf('  %s %s -> %s\n', tag, who, heading);

        % save progress after every student so an interruption loses nothing
        if ~cfg.dryRun
            g.upload(processedPath, ba_utf8_bytes(strjoin([processed {''}], newline)));
        end
    catch ME
        nFailed = nFailed + 1;
        fprintf(2, '  FAILED %s: %s\n', who, ME.message);
        reviewTxt = [reviewTxt ba_csv_line({runTime, key, who, ['FAILED: ' ME.message]})]; %#ok<AGROW>
    end
end

%% 4. Save logs
if ~cfg.dryRun
    g.upload(processedPath, ba_utf8_bytes(strjoin([processed {''}], newline)));
    g.upload(reviewPath, ba_utf8_bytes(reviewTxt));
end

fprintf('\nDone. Updated: %d | Skipped (see review_log.csv): %d | Failed: %d | Already done before: %d\n', ...
        nDone, nSkipped, nFailed, nOld);

%% ======================= local helper functions =======================

function c = ba_find_column(T, patterns, exact)
    if nargin < 3, exact = false; end
    names = T.Properties.VariableNames;
    c = '';
    for p = patterns
        if exact
            hit = find(strcmpi(strtrim(names), p{1}), 1);
        else
            hit = find(contains(names, p{1}, 'IgnoreCase', true), 1);
        end
        if ~isempty(hit), c = names{hit}; return; end
    end
end

function s = ba_cell_text(T, col, i)
    if isempty(col), s = ''; return; end
    x = T.(col)(i);
    if iscell(x), x = x{1}; end
    if ismissing(x), s = ''; return; end
    if isnumeric(x)
        if x == round(x), s = sprintf('%.0f', x); else, s = num2str(x); end
    elseif isdatetime(x)
        s = char(x, 'yyyy-MM-dd HH:mm');
    else
        s = char(string(x));
    end
    s = strtrim(s);
end

function n = ba_clean_name(n)
    % "  hanadi   SALEH " -> "Hanadi Saleh" (so case differences don't create
    % a second file for the same student)
    n = regexprep(strtrim(n), '\s+', ' ');
    n = regexprep(n, '[\\/:*?"<>|.,;()\[\]{}#%&]', '');   % characters not allowed/awkward in file names
    words = strsplit(lower(n), ' ');
    for k = 1:numel(words)
        if ~isempty(words{k}), words{k}(1) = upper(words{k}(1)); end
    end
    n = strjoin(words, ' ');
end

function tf = ba_is_agreement(s)
    s = lower(strtrim(s));
    tf = startsWith(s, 'yes') || (contains(s, 'agree') && ~contains(s, 'not agree') && ~contains(s, 'disagree'));
end

function h = ba_category_heading(answer)
    % Map the Forms category answer to the heading printed on the transcript.
    a = lower(answer);
    if contains(a, 'community')
        h = 'COMMUNITY IMPACT';
    elseif contains(a, 'competition')
        h = 'NATIONAL & INTERNATIONAL COMPETITIONS';
    elseif contains(a, 'professional')
        h = 'PROFESSIONAL EXPERIENCE';
    elseif contains(a, 'research')
        h = 'RESEARCH';
    elseif contains(a, 'global') || contains(a, 'exchange')
        h = 'GLOBAL EDUCATION & EXCHANGE';
    elseif contains(a, 'award') || contains(a, 'recognition')
        h = 'AWARDS & RECOGNITIONS';
    else
        h = '';
    end
end

function line = ba_csv_line(fields)
    for k = 1:numel(fields)
        fields{k} = ['"' strrep(fields{k}, '"', '""') '"'];
    end
    line = [strjoin(fields, ',') newline];
end

function lines = ba_read_lines(bytes)
    txt = ba_utf8(bytes);
    lines = strtrim(strsplit(txt, {'\r\n', '\n'}));
    lines = lines(~cellfun(@isempty, lines));
end

function t = ba_utf8(bytes)
    if isempty(bytes), t = ''; else, t = native2unicode(uint8(bytes(:)'), 'UTF-8'); end
end

function b = ba_utf8_bytes(t)
    b = unicode2native(t, 'UTF-8');
end

function ba_write_bytes(p, bytes)
    fid = fopen(p, 'w'); fwrite(fid, bytes, 'uint8'); fclose(fid);
end

function b = ba_read_bytes(p)
    fid = fopen(p, 'r'); b = fread(fid, Inf, '*uint8')'; fclose(fid);
end

function ba_safe_rmdir(d)
    try
        if isfolder(d), rmdir(d, 's'); end
    catch
    end
end

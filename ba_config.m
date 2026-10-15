function CFG = ba_config()
%BA_CONFIG Settings for the Beyond Academics transcript automation.
%
%   This is the only file you should need to edit.

%% Azure / Microsoft Entra app registration (provided by IT)
% Neither value is a secret: they only identify the app and the AUB tenant.
CFG.tenantId = 'PASTE-DIRECTORY-TENANT-ID-HERE';
CFG.clientId = 'PASTE-APPLICATION-CLIENT-ID-HERE';

% Delegated permissions requested at sign-in.
%   Files.ReadWrite  -> the signed-in account's own OneDrive files
%   offline_access   -> lets the script refresh its sign-in without asking again
CFG.scopes = 'Files.ReadWrite offline_access';

%% OneDrive locations
% Paths are relative to the root of the OneDrive of the account you SIGN IN
% AS (the account that owns the Beyond Academics form and its files).
CFG.responsesPath = 'Beyond Academics/FormResponses.xlsx';
CFG.templatePath  = 'Beyond Academics Transcripts/_automation/BEYOND_ACADEMICS_Transcript_TEMPLATE.docx';

% Where student transcripts are written. Point this at a TEST folder first.
CFG.outputFolder  = 'Beyond Academics Transcripts/TEST';

% Where the run logs are kept (processed_log.csv, review_log.csv).
CFG.stateFolder   = 'Beyond Academics Transcripts/_automation';

%% Local settings
% Sign-in token cache. Kept in MATLAB's private preferences folder, never in
% the code folder, so it can't be committed to git by accident.
CFG.tokenCache = fullfile(prefdir, 'beyond_academics_graph_token.mat');

% Format of "Date Issued" on the transcript.
CFG.dateFormat = 'dd-MM-yyyy';

% Each category box on the template shows about this many rows. Students
% with more entries in one category are flagged in review_log.csv.
CFG.visibleRowsPerCategory = 3;

% true = do everything except upload (prints what it would change).
CFG.dryRun = false;
end

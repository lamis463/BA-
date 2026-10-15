classdef TranscriptDocx
    %TRANSCRIPTDOCX Edits the official Beyond Academics transcript (.docx)
    %   directly as XML, without needing Microsoft Word.
    %
    %   A .docx file is a ZIP archive; the visible content is in
    %   word/document.xml. On the official template, each of the six
    %   category tables sits inside its own text box and starts with:
    %     row 1: the category heading (e.g. COMMUNITY IMPACT)
    %     row 2: column labels SEMESTER | TITLE | HOURS/DAYS
    %     rows 3+: entries (the template ships with 3 blank rows)
    %
    %   Nothing else in the document is touched.

    methods (Static)

        %% ---------- file level ----------

        function xml = readXml(docxPath, workDir)
            % Unzips docxPath into workDir and returns word/document.xml.
            if isfolder(workDir), rmdir(workDir, 's'); end
            mkdir(workDir);
            unzip(docxPath, workDir);
            xml = TranscriptDocx.readUtf8(fullfile(workDir, 'word', 'document.xml'));
        end

        function writeDocx(workDir, xml, outPath)
            % Writes xml back into workDir and zips it into outPath.
            TranscriptDocx.writeUtf8(fullfile(workDir, 'word', 'document.xml'), xml);
            items = dir(workDir);
            names = setdiff({items.name}, {'.', '..'});
            tmpZip = [tempname '.zip'];
            % Third argument = root folder, so entries are stored relative to
            % it ([Content_Types].xml must sit at the root of the archive).
            zip(tmpZip, names, workDir);
            movefile(tmpZip, outPath, 'f');
        end

        %% ---------- header ----------

        function xml = setHeader(xml, studentName, concentration, dateStr)
            % Student name and concentration are filled only if still blank,
            % so a later submission never overwrites them. Date Issued is
            % refreshed on every update.
            nameRun = TranscriptDocx.makeRun([' ' studentName], '');
            xml = TranscriptDocx.fillAfterLabel(xml, '<w:t>name:</w:t></w:r>', nameRun);

            if ~isempty(strtrim(concentration))
                concRun = TranscriptDocx.makeRun([' ' concentration], '<w:spacing w:val="-2"/>');
                xml = TranscriptDocx.fillAfterLabel(xml, '<w:t>Concentration(s):</w:t></w:r>', concRun);
            end

            xml = regexprep(xml, 'Date Issued: [0-9]{2}[-/.][0-9]{2}[-/.][0-9]{4}', ...
                            ['Date Issued: ' dateStr], 'once');
        end

        %% ---------- category tables ----------

        function [xml, nEntries] = addEntry(xml, heading, semester, title, hours)
            % Puts one entry into the category table whose heading matches.
            % Uses the first blank row; if none is left, adds a new row.
            % nEntries = number of filled entry rows in that table afterwards.
            [ts, te] = regexp(xml, '<w:tbl>.*?</w:tbl>', 'start', 'end');
            target = TranscriptDocx.normalise(heading);
            k = 0;
            for i = 1:numel(ts)
                tbl = xml(ts(i):te(i));
                firstRow = regexp(tbl, '<w:tr[ >].*?</w:tr>', 'match', 'once');
                if strcmp(TranscriptDocx.normalise(TranscriptDocx.textOf(firstRow)), target)
                    k = i;
                    break;
                end
            end
            if k == 0
                error('TranscriptDocx:noTable', 'No table with heading "%s" found in the transcript.', heading);
            end

            tbl = xml(ts(k):te(k));
            [rs, re] = regexp(tbl, '<w:tr[ >].*?</w:tr>', 'start', 'end');
            if numel(rs) < 2
                error('TranscriptDocx:badTable', 'Table "%s" has an unexpected layout.', heading);
            end

            values = {semester, title, hours};
            filled = 0;
            blankIdx = 0;
            for r = 3:numel(rs)
                rowXml = tbl(rs(r):re(r));
                if isempty(regexp(rowXml, '<w:t[ >]', 'once'))
                    if blankIdx == 0, blankIdx = r; end
                else
                    filled = filled + 1;
                end
            end

            if blankIdx > 0
                newRow = TranscriptDocx.fillRow(tbl(rs(blankIdx):re(blankIdx)), values);
                tbl = [tbl(1:rs(blankIdx)-1) newRow tbl(re(blankIdx)+1:end)];
            else
                % No blank rows left: copy the last row's layout and append it.
                proto = tbl(rs(end):re(end));
                proto = regexprep(proto, ' w14:(paraId|textId)="[^"]*"', '');
                newRow = TranscriptDocx.fillRow(proto, values);
                tbl = [tbl(1:re(end)) newRow tbl(re(end)+1:end)];
            end
            nEntries = filled + 1;
            xml = [xml(1:ts(k)-1) tbl xml(te(k)+1:end)];
        end

        function n = countEntries(xml, heading)
            % Number of filled entry rows in a category table.
            [ts, te] = regexp(xml, '<w:tbl>.*?</w:tbl>', 'start', 'end');
            n = 0;
            for i = 1:numel(ts)
                tbl = xml(ts(i):te(i));
                rows = regexp(tbl, '<w:tr[ >].*?</w:tr>', 'match');
                if ~isempty(rows) && strcmp(TranscriptDocx.normalise(TranscriptDocx.textOf(rows{1})), ...
                                            TranscriptDocx.normalise(heading))
                    for r = 3:numel(rows)
                        if ~isempty(regexp(rows{r}, '<w:t[ >]', 'once')), n = n + 1; end
                    end
                    return;
                end
            end
        end

        %% ---------- helpers ----------

        function s = escapeXml(s)
            s = strrep(s, '&', '&amp;');
            s = strrep(s, '<', '&lt;');
            s = strrep(s, '>', '&gt;');
            s = strrep(s, '"', '&quot;');
            % drop control characters that are illegal in XML
            s = regexprep(s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '');
        end

        function t = textOf(fragment)
            parts = regexp(fragment, '<w:t(?: [^>]*)?>([^<]*)</w:t>', 'tokens');
            t = '';
            for i = 1:numel(parts), t = [t parts{i}{1}]; end %#ok<AGROW>
        end
    end

    methods (Static, Access = private)
        function xml = fillAfterLabel(xml, labelRunEnd, newRun)
            % Insert newRun right after the label, only if nothing follows it
            % yet in that paragraph (i.e. the field is still blank).
            blankPattern = [labelRunEnd '</w:p>'];
            idx = strfind(xml, blankPattern);
            if ~isempty(idx)
                p = idx(1) + numel(labelRunEnd) - 1;
                xml = [xml(1:p) newRun xml(p+1:end)];
            end
        end

        function r = makeRun(text, extraRPr)
            r = ['<w:r><w:rPr>' extraRPr '</w:rPr><w:t xml:space="preserve">' ...
                 TranscriptDocx.escapeXml(text) '</w:t></w:r>'];
        end

        function rowXml = fillRow(rowXml, values)
            % Replace the paragraph in each of the 3 cells with one that holds
            % the value, styled like the template's SEMESTER/TITLE labels.
            [cs, ce] = regexp(rowXml, '<w:tc>.*?</w:tc>', 'start', 'end');
            if numel(cs) ~= numel(values)
                error('TranscriptDocx:badRow', 'Expected %d cells in an entry row, found %d.', numel(values), numel(cs));
            end
            for c = numel(cs):-1:1          % back to front so indices stay valid
                cellXml = rowXml(cs(c):ce(c));
                [ps, pe] = regexp(cellXml, '<w:p[ >].*?</w:p>', 'start', 'end', 'once');
                para = ['<w:p><w:pPr><w:pStyle w:val="TableParagraph"/>' ...
                        '<w:spacing w:before="64" w:after="40"/><w:ind w:left="115" w:right="60"/>' ...
                        '<w:rPr><w:sz w:val="14"/></w:rPr></w:pPr>' ...
                        '<w:r><w:rPr><w:sz w:val="14"/></w:rPr><w:t xml:space="preserve">' ...
                        TranscriptDocx.escapeXml(values{c}) '</w:t></w:r></w:p>'];
                cellXml = [cellXml(1:ps-1) para cellXml(pe+1:end)];
                rowXml = [rowXml(1:cs(c)-1) cellXml rowXml(ce(c)+1:end)];
            end
        end

        function s = normalise(s)
            s = strrep(s, '&amp;', '&');
            s = upper(regexprep(strtrim(s), '\s+', ' '));
        end

        function s = readUtf8(p)
            fid = fopen(p, 'r', 'n', 'UTF-8');
            c = onCleanup(@() fclose(fid));
            s = fread(fid, [1 Inf], '*char');
        end

        function writeUtf8(p, s)
            fid = fopen(p, 'w', 'n', 'UTF-8');
            c = onCleanup(@() fclose(fid));
            fwrite(fid, s, 'char');
        end
    end
end

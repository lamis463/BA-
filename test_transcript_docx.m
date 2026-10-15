function test_transcript_docx(templatePath, outPath)
%TEST_TRANSCRIPT_DOCX Offline check of the .docx editing (no OneDrive needed).
%   test_transcript_docx('BEYOND_ACADEMICS_Transcript_TEMPLATE.docx', 'sample.docx')
%   Fills a copy of the template with sample entries; open sample.docx in Word.
work = fullfile(tempdir, 'ba_test');
xml = TranscriptDocx.readXml(templatePath, work);
xml = TranscriptDocx.setHeader(xml, 'Sample Student', 'Finance', '14-10-2026');
xml = TranscriptDocx.setHeader(xml, 'MUST NOT OVERWRITE', 'MUST NOT OVERWRITE', '15-10-2026');
[xml, n1] = TranscriptDocx.addEntry(xml, 'COMMUNITY IMPACT', 'Fall 2025', 'Red Cross volunteer', '40 hours');
[xml, n2] = TranscriptDocx.addEntry(xml, 'COMMUNITY IMPACT', 'Spring 2026', 'Beach cleanup <Byblos> & "Tyre"', '2 days');
[xml, n3] = TranscriptDocx.addEntry(xml, 'AWARDS & RECOGNITIONS', 'Spring 2026', 'Dean''s Honor List', '-');
assert(n1 == 1 && n2 == 2 && n3 == 1, 'entry counts wrong');
assert(TranscriptDocx.countEntries(xml, 'COMMUNITY IMPACT') == 2);
assert(isempty(strfind(xml, 'MUST NOT OVERWRITE')), 'header was overwritten');
assert(~isempty(strfind(xml, 'Date Issued: 15-10-2026')), 'date not refreshed');
TranscriptDocx.writeDocx(work, xml, outPath);
fprintf('OK - wrote %s\n', outPath);
end

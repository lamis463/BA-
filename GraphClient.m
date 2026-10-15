classdef GraphClient < handle
    %GRAPHCLIENT Minimal Microsoft Graph client for reading and writing
    %   OneDrive files as the signed-in user (delegated permissions).
    %
    %   Sign-in uses the OAuth 2.0 device code flow: the script prints a
    %   short code, the user opens https://microsoft.com/devicelogin in any
    %   browser, enters it and signs in (MFA included). No password ever
    %   passes through this code. The refresh token is cached locally so
    %   later runs don't need to sign in again until it expires.
    %
    %   Only these Graph endpoints are used:
    %     GET  /me/drive/root:/{path}            (file metadata + download URL)
    %     PUT  /me/drive/root:/{path}:/content   (upload/replace a file < 4 MB)

    properties (SetAccess = private)
        TenantId
        ClientId
        Scopes
        CachePath
    end

    properties (Access = private)
        AccessToken  = ''
        RefreshToken = ''
        ExpiresAt    = 0      % posix seconds
    end

    properties (Constant)
        GraphBase = 'https://graph.microsoft.com/v1.0'
    end

    methods
        function obj = GraphClient(tenantId, clientId, scopes, cachePath)
            obj.TenantId  = tenantId;
            obj.ClientId  = clientId;
            obj.Scopes    = scopes;
            obj.CachePath = cachePath;
        end

        function login(obj)
            % Reuse a cached sign-in if there is one; otherwise device code.
            if isfile(obj.CachePath)
                c = load(obj.CachePath);
                if isfield(c, 'refreshToken') && ~isempty(c.refreshToken)
                    obj.RefreshToken = c.refreshToken;
                    if obj.refresh()
                        fprintf('Signed in (cached session).\n');
                        return;
                    end
                end
            end
            obj.deviceCodeLogin();
        end

        function logout(obj)
            obj.AccessToken = ''; obj.RefreshToken = ''; obj.ExpiresAt = 0;
            if isfile(obj.CachePath), delete(obj.CachePath); end
        end

        function bytes = download(obj, drivePath)
            % Returns file contents as uint8, or [] if the file doesn't exist.
            url = [obj.GraphBase '/me/drive/root:/' GraphClient.encodePath(drivePath)];
            [status, txt] = obj.request('GET', url, [], '');
            if status == 404
                bytes = [];
                return;
            end
            GraphClient.check(status, txt, ['read metadata of ' drivePath]);
            meta = jsondecode(txt);
            f = fieldnames(meta);
            k = find(~cellfun(@isempty, strfind(f, 'downloadUrl')), 1); %#ok<STRCLFH>
            if isempty(k)
                error('GraphClient:noDownloadUrl', 'No download URL for %s (is it a folder?)', drivePath);
            end
            % The download URL is pre-authorised and short-lived: no token sent.
            [status, bytes] = GraphClient.rawRequest('GET', meta.(f{k}), {}, []);
            GraphClient.check(status, GraphClient.utf8(bytes), ['download ' drivePath]);
            bytes = uint8(bytes(:)');
        end

        function upload(obj, drivePath, bytes)
            % Creates or replaces a file (simple upload, files under 4 MB).
            if numel(bytes) > 4e6
                error('GraphClient:tooLarge', '%s is over 4 MB; simple upload not supported.', drivePath);
            end
            url = [obj.GraphBase '/me/drive/root:/' GraphClient.encodePath(drivePath) ':/content'];
            [status, txt] = obj.request('PUT', url, uint8(bytes), 'application/octet-stream');
            GraphClient.check(status, txt, ['upload ' drivePath]);
        end
    end

    methods (Access = private)
        function deviceCodeLogin(obj)
            url = sprintf('https://login.microsoftonline.com/%s/oauth2/v2.0/devicecode', obj.TenantId);
            [status, txt] = GraphClient.postForm(url, {'client_id', obj.ClientId, 'scope', obj.Scopes});
            GraphClient.check(status, txt, 'start sign-in');
            dc = jsondecode(txt);

            fprintf('\n=============================================================\n');
            fprintf('%s\n', dc.message);
            fprintf('=============================================================\n\n');

            tokenUrl = obj.tokenUrl();
            deadline = now + dc.expires_in / 86400;
            interval = dc.interval;
            while now < deadline
                pause(interval);
                [status, txt] = GraphClient.postForm(tokenUrl, { ...
                    'grant_type',  'urn:ietf:params:oauth:grant-type:device_code', ...
                    'client_id',   obj.ClientId, ...
                    'device_code', dc.device_code});
                r = jsondecode(txt);
                if status == 200
                    obj.storeTokens(r);
                    fprintf('Signed in.\n');
                    return;
                end
                switch r.error
                    case 'authorization_pending'
                        % keep waiting
                    case 'slow_down'
                        interval = interval + 5;
                    otherwise
                        error('GraphClient:login', 'Sign-in failed: %s - %s', r.error, r.error_description);
                end
            end
            error('GraphClient:login', 'Sign-in code expired. Run the script again.');
        end

        function ok = refresh(obj)
            [status, txt] = GraphClient.postForm(obj.tokenUrl(), { ...
                'grant_type',    'refresh_token', ...
                'client_id',     obj.ClientId, ...
                'refresh_token', obj.RefreshToken, ...
                'scope',         obj.Scopes});
            ok = (status == 200);
            if ok
                obj.storeTokens(jsondecode(txt));
            end
        end

        function storeTokens(obj, r)
            obj.AccessToken = r.access_token;
            obj.ExpiresAt   = GraphClient.nowPosix() + double(r.expires_in);
            if isfield(r, 'refresh_token')
                obj.RefreshToken = r.refresh_token;
            end
            refreshToken = obj.RefreshToken; %#ok<NASGU>
            save(obj.CachePath, 'refreshToken');
        end

        function tok = token(obj)
            if isempty(obj.AccessToken) || GraphClient.nowPosix() > obj.ExpiresAt - 300
                if isempty(obj.RefreshToken) || ~obj.refresh()
                    obj.deviceCodeLogin();
                end
            end
            tok = obj.AccessToken;
        end

        function u = tokenUrl(obj)
            u = sprintf('https://login.microsoftonline.com/%s/oauth2/v2.0/token', obj.TenantId);
        end

        function [status, txt] = request(obj, method, url, body, contentType)
            % Authorised Graph call with simple retry on throttling / server errors.
            for attempt = 1:5
                headers = {'Authorization', ['Bearer ' obj.token()]};
                if ~isempty(contentType)
                    headers(end+1, :) = {'Content-Type', contentType}; %#ok<AGROW>
                end
                [status, raw, retryAfter] = GraphClient.rawRequest(method, url, headers, body);
                txt = GraphClient.utf8(raw);
                if ~(status == 429 || status >= 500) || attempt == 5
                    return;
                end
                pause(max(retryAfter, 2 * attempt));
            end
        end
    end

    methods (Static)
        function p = encodePath(drivePath)
            % Percent-encode each path segment (spaces, &, Arabic letters...).
            parts = strsplit(strrep(drivePath, '\', '/'), '/');
            parts = parts(~cellfun(@isempty, parts));
            for i = 1:numel(parts)
                b = unicode2native(parts{i}, 'UTF-8');
                out = '';
                for c = b
                    ch = char(c);
                    if (c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || any(ch == '-_.~')
                        out = [out ch]; %#ok<AGROW>
                    else
                        out = [out sprintf('%%%02X', c)]; %#ok<AGROW>
                    end
                end
                parts{i} = out;
            end
            p = strjoin(parts, '/');
        end
    end

    methods (Static, Access = private)
        function [status, txt] = postForm(url, pairs)
            body = '';
            for i = 1:2:numel(pairs)
                if i > 1, body = [body '&']; end %#ok<AGROW>
                body = [body pairs{i} '=' GraphClient.formEncode(pairs{i+1})]; %#ok<AGROW>
            end
            [status, raw] = GraphClient.rawRequest('POST', url, ...
                {'Content-Type', 'application/x-www-form-urlencoded'}, uint8(body));
            txt = GraphClient.utf8(raw);
        end

        function s = formEncode(v)
            b = unicode2native(char(v), 'UTF-8');
            s = '';
            for c = b
                ch = char(c);
                if (c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || any(ch == '-_.~')
                    s = [s ch]; %#ok<AGROW>
                else
                    s = [s sprintf('%%%02X', c)]; %#ok<AGROW>
                end
            end
        end

        function [status, raw, retryAfter] = rawRequest(method, url, headers, body)
            % One HTTP call. Never throws on 4xx/5xx; returns the status.
            import matlab.net.http.*
            import matlab.net.http.field.*
            fields = matlab.net.http.HeaderField.empty;
            ctype = '';
            for i = 1:size(headers, 1)
                if strcmpi(headers{i, 1}, 'Content-Type')
                    ctype = headers{i, 2};
                    fields(end+1) = ContentTypeField(ctype); %#ok<AGROW>
                else
                    fields(end+1) = HeaderField(headers{i, 1}, headers{i, 2}); %#ok<AGROW>
                end
            end
            if isempty(body)
                msgBody = [];
                if strcmp(method, 'PUT') || strcmp(method, 'POST')
                    msgBody = MessageBody(uint8([]));
                end
            else
                msgBody = MessageBody();
                msgBody.Payload = uint8(body(:)');   % send bytes exactly as given
            end
            req = RequestMessage(method, fields, msgBody);
            opts = HTTPOptions('ConvertResponse', false, 'SavePayload', true, 'ConnectTimeout', 30);
            resp = req.send(matlab.net.URI(url), opts);
            status = double(resp.StatusCode);
            raw = uint8([]);
            if ~isempty(resp.Body) && ~isempty(resp.Body.Payload)
                raw = uint8(resp.Body.Payload);
            end
            retryAfter = 0;
            ra = resp.getFields('Retry-After');
            if ~isempty(ra)
                retryAfter = str2double(ra(1).Value);
                if isnan(retryAfter), retryAfter = 0; end
            end
        end

        function t = utf8(raw)
            if isempty(raw), t = ''; else, t = native2unicode(uint8(raw(:)'), 'UTF-8'); end
        end

        function check(status, txt, what)
            if status < 200 || status >= 300
                if numel(txt) > 600, txt = [txt(1:600) '...']; end
                error('GraphClient:http', 'Could not %s (HTTP %d): %s', what, status, txt);
            end
        end

        function t = nowPosix()
            t = posixtime(datetime('now', 'TimeZone', 'UTC'));
        end
    end
end

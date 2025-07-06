unit TeamsGraphDemo;

interface

uses
  System.SysUtils, System.Classes, System.JSON, System.Net.HttpClient, System.Net.URLClient, System.Net.HttpClientComponent;

type
  TTeamsConnector = class
  private
    FClientId: string;
    FClientSecret: string;
    FTenantId: string;
    FAccessToken: string;
    function GetToken: Boolean;
  public
    constructor Create(const ClientId, ClientSecret, TenantId: string);
    function SendMessageToChannel(const TeamId, ChannelId, MessageText: string): Boolean;
    property AccessToken: string read FAccessToken;
  end;

implementation

const
  TOKEN_URL = 'https://login.microsoftonline.com/%s/oauth2/v2.0/token';
  GRAPH_URL = 'https://graph.microsoft.com/v1.0';

constructor TTeamsConnector.Create(const ClientId, ClientSecret, TenantId: string);
begin
  FClientId := ClientId;
  FClientSecret := ClientSecret;
  FTenantId := TenantId;
  FAccessToken := '';
end;

// Получение токена OAuth2 client credentials
function TTeamsConnector.GetToken: Boolean;
var
  Http: TNetHTTPClient;
  Params: string;
  Resp: IHTTPResponse;
  Json: TJSONObject;
  Content: TStringStream;
begin
  Result := False;
  Http := TNetHTTPClient.Create(nil);
  try
    Params :=
      'client_id=' + FClientId +
      '&scope=https://graph.microsoft.com/.default' +
      '&client_secret=' + FClientSecret +
      '&grant_type=client_credentials';
    Content := TStringStream.Create(Params, TEncoding.UTF8);
    try
      Http.ContentType := 'application/x-www-form-urlencoded';
      Resp := Http.Post(Format(TOKEN_URL, [FTenantId]), Content);
      if Resp.StatusCode = 200 then
      begin
        Json := TJSONObject.ParseJSONValue(Resp.ContentAsString) as TJSONObject;
        try
          if Json.TryGetValue('access_token', FAccessToken) then
            Result := True;
        finally
          Json.Free;
        end;
      end;
    finally
      Content.Free;
    end;
  finally
    Http.Free;
  end;
end;

// Отправить сообщение в канал Teams
function TTeamsConnector.SendMessageToChannel(const TeamId, ChannelId, MessageText: string): Boolean;
var
  Http: TNetHTTPClient;
  ReqBody: TJSONObject;
  Resp: IHTTPResponse;
begin
  Result := False;
  if (FAccessToken = '') and not GetToken then Exit;

  Http := TNetHTTPClient.Create(nil);
  try
    Http.CustomHeaders['Authorization'] := 'Bearer ' + FAccessToken;
    Http.ContentType := 'application/json';
    ReqBody := TJSONObject.Create;
    try
      ReqBody.AddPair('body', TJSONObject.Create.AddPair('content', MessageText));
      Resp := Http.Post(
        GRAPH_URL + Format('/teams/%s/channels/%s/messages', [TeamId, ChannelId]),
        TStringStream.Create(ReqBody.ToJSON, TEncoding.UTF8)
      );
      Result := Resp.StatusCode = 201;
    finally
      ReqBody.Free;
    end;
  finally
    Http.Free;
  end;
end;

end.

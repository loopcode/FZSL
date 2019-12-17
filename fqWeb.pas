unit fqWeb;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, Registry, SHDocVw;

type
  TfqWebForm = class(TForm)
    WebBrowser1: TWebBrowser;
  private
    { Private declarations }
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    procedure LoadURL(URL: string; DB: string = '');
  end;

implementation

{$R *.dfm}

constructor TfqWebForm.Create(AOwner: TComponent);
var
  reg: TRegistry;
begin
  // ���� IE �ں�
  // ϵͳ IE �ں˸��� IE7 ʱ��TWebBrowser �ؼ�Ĭ�ϴ������� IE7 ����ģʽ��Ӧ�ó������ͨ�����·����޸ļ��ݰ汾
  // HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION
  // ���һ������ΪӦ�ó����������Ƶ� DWORD ֵ��IE7/7000��IE8/8000��IE9/9000��IE10/10000��IE11/11000��Edge/>12000(ֻ��ģ�� userAgent���ں���Ȼ�� IE11)
  //
  // ���� UAC ����
  // �� Vista ��ʼ������ϵͳ������ UAC ���ƹ��ܣ�������� UAC ������£�û��ʹ�ù���ԱȨ��ִ�г���Ĭ�϶������ UAC ���⻯
  // ��ע���û��д��Ȩ��ʱ���������ᱻ�Զ��ض��� HKEY_CURRENT_USER\Software\Classes\VirtualStore\MACHINE\...
  // ��Ŀ¼û��д��Ȩ��ʱ���������ᱻ�Զ��ض��� C:\Users\�û���\AppData\Local\VirtualStore\...
  // ������ϲ�����û��ʹ�ù���ԱȨ��ִ�г���ʱ�ᱻ�ض�������λ��
  // HKEY_CURRENT_USER\Software\Classes\VirtualStore\MACHINE\SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION

  reg := TRegistry.Create();
  reg.RootKey := HKEY_LOCAL_MACHINE;
  reg.OpenKey('SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION', True);
  reg.WriteInteger(ExtractFileName(Application.ExeName), 11000);
  reg.CloseKey();
  reg.Free();

  inherited Create(AOwner);
end;

procedure TfqWebForm.LoadURL(URL: string; DB: string);
var
  FileName: string;
  Stream: TStream;
begin
  if (DB <> '') then
  begin
    FileName := ExtractFilePath(Application.ExeName) + 'Web\database.js';
    Stream := TFileStream.Create(FileName, fmCreate);
    DB := 'var dbFile = "' + StringReplace(DB, '\', '\\', [rfReplaceAll	]) + '";';
    DB := UTF8Encode(WideString(DB));
    try
      Stream.WriteBuffer(Pointer(DB)^, Length(DB));
    finally
      Stream.Free();
    end;
  end;

  WebBrowser1.Navigate(WideString(URL));
end;

end.

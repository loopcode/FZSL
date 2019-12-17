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
  // 关于 IE 内核
  // 系统 IE 内核高于 IE7 时，TWebBrowser 控件默认创建的是 IE7 兼容模式，应用程序可以通过如下方法修改兼容版本
  // HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION
  // 添加一个名称为应用程序自身名称的 DWORD 值，IE7/7000，IE8/8000，IE9/9000，IE10/10000，IE11/11000，Edge/>12000(只是模拟 userAgent，内核仍然是 IE11)
  //
  // 关于 UAC 控制
  // 自 Vista 开始，操作系统增加了 UAC 控制功能，如果开启 UAC 的情况下，没有使用管理员权限执行程序，默认都会采用 UAC 虚拟化
  // 对注册表没有写入权限时，操作将会被自动重定向到 HKEY_CURRENT_USER\Software\Classes\VirtualStore\MACHINE\...
  // 对目录没有写入权限时，操作将会被自动重定向到 C:\Users\用户名\AppData\Local\VirtualStore\...
  // 因此以上操作在没有使用管理员权限执行程序时会被重定向到如下位置
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

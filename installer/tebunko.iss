; tebunko のインストーラー（Inno Setup 7）
;
; ビルドは tools\new_installer.ps1 が行う（release.yml からも呼ぶ）。次の値をコマンドラインの /D で渡す:
;   AppVersion      表示する版（例 0.2.0。アンインストールの一覧に出る）
;   SetupVersion    インストーラーのファイル名に付ける版（タグの名前。例 v0.2.0。zip の名前と揃える）
;   NumericVersion  exe の版の情報に書く数字だけの版（例 0.2.0.0）
;   StageDir        入れるファイルを並べたフォルダ（tebunko.exe・scripts\・LICENSE）
;   IconFile        インストーラーのアイコン（scripts\tebunko\tebunko.ico）
;
; 方針（docs/safety/disclosure.md「インストーラー版」）:
;   ・管理者権限なしで、利用者ごとの %LOCALAPPDATA%\Programs\tebunko に入れる。管理者なら Program Files も選べる
;   ・入れるのは tebunko.exe（installer\tebunko.cs）・scripts\・LICENSE だけ。レジストリに書くのは、Windows のインストーラーが
;     必ず書くアンインストールの情報だけ（[Registry] は使わない）。サービス・自動起動・PATH・ファイルの関連付けは触らない
;   ・更新は、新しい版のインストーラーを実行するだけにする。古い版で消したスクリプトが残らないよう、scripts\ を消してから入れる
;   ・アンインストールでは、入れたファイルと setting.config を消す。インデックス（ワークスペース）は利用者が選んだ場所にあり得るため
;     消さずに、残っている場所を知らせる

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef SetupVersion
  #define SetupVersion AppVersion
#endif
#ifndef NumericVersion
  #define NumericVersion "0.0.0.0"
#endif
#ifndef StageDir
  #error StageDir を /D で指定してください（tools\new_installer.ps1 から実行する）
#endif
#ifndef IconFile
  #error IconFile を /D で指定してください（tools\new_installer.ps1 から実行する）
#endif

[Setup]
; AppId はアンインストールの情報の鍵。版をまたいで同じにする（変えると、別のアプリとして並んで入る）
AppId={{E2FA5AC9-5A36-41AF-B5E2-B989E2878D3C}
AppName=tebunko
AppVersion={#AppVersion}
AppVerName=tebunko {#AppVersion}
AppPublisher=hsgwa
AppPublisherURL=https://github.com/hsgwa/tebunko
AppSupportURL=https://github.com/hsgwa/tebunko/issues
AppUpdatesURL=https://github.com/hsgwa/tebunko/releases
VersionInfoVersion={#NumericVersion}
VersionInfoProductName=tebunko
VersionInfoDescription=tebunko のインストーラー
; 既定は管理者権限なし（%LOCALAPPDATA%\Programs）。管理者は、最初の画面で「すべてのユーザー」（Program Files）も選べる
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
DefaultDirName={autopf}\tebunko
DisableProgramGroupPage=yes
; 64 ビットの Windows では、管理者のときの入れ先を Program Files（x86 ではない方）にする。スクリプトと exe は CPU を問わない
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
; tebunko.exe の起動中はインストール・アンインストールを止め、閉じるよう求める（tebunko.cs のミューテックス）
AppMutex=tebunko-installed-app
CloseApplications=no
SetupIconFile={#IconFile}
UninstallDisplayIcon={app}\tebunko.exe
; アンインストーラー（Inno Setup が unins000.exe と名付ける。名前を変える設定は無い）は、tebunko のものと分かるよう uninstall\ に置く
UninstallFilesDir={app}\uninstall
UninstallDisplayName=tebunko
WizardStyle=modern
Compression=lzma2/max
SolidCompression=yes
OutputBaseFilename=tebunko-setup-{#SetupVersion}

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作る"; Flags: unchecked

[InstallDelete]
; 前の版で入れた scripts\ を消してから入れる（前の版にだけあったスクリプトを残さない）
Type: filesandordirs; Name: "{app}\scripts"

[Files]
Source: "{#StageDir}\tebunko.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#StageDir}\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#StageDir}\scripts\*"; DestDir: "{app}\scripts"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\tebunko"; Filename: "{app}\tebunko.exe"
Name: "{autodesktop}\tebunko"; Filename: "{app}\tebunko.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\tebunko.exe"; Description: "tebunko を起動する"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 利用者ごとの設定（ツールのフォルダに書き込めるときはここに作る。scripts\shared\core\data_dir.ps1）
Type: files; Name: "{app}\setting.config"

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  // インデックスには文書の文字がそのまま入っているため、残っている場所を知らせる（無人のアンインストールでは出さない）
  if CurUninstallStep = usPostUninstall then
    SuppressibleMsgBox(
      'tebunko を削除しました。' + #13#10#13#10 +
      'インデックス・取り込み一覧・ログ（ワークスペース）は削除していません。' +
      '文書の文字がそのまま入っているため、不要なら次のフォルダを削除してください。' + #13#10#13#10 +
      '・ワークスペース（既定は ' + ExpandConstant('{%USERPROFILE}') + '\Documents\tebunko_ws。［8 設定］で変えた場合はそのフォルダ）' + #13#10 +
      '・' + ExpandConstant('{localappdata}') + '\tebunko（Program Files に入れた場合の設定など）',
      mbInformation, MB_OK, IDOK);
end;

### <a name="configuration"></a>構成

次のファイルは、アドインの構成設定を指定します。

- プロジェクトのルート ディレクトリにある **./manifest.xml** ファイルで、アドインの機能と設定を定義します。

- **./.プロジェクト** のルート ディレクトリの ENV ファイルは、アドイン プロジェクトで使用される定数を定義します。

### <a name="task-pane"></a>作業ウィンドウ 

次のファイルは、アドインの作業ウィンドウの UI と機能を定義します。

- **./src/taskpane/taskpane.html** ファイルには、作業ペイン用のHTMLマークアップが含まれています。

- **./src/taskpane/taskpane.css** ファイルには、作業ペインのコンテンツに適用されるCSSが含まれています。

- JavaScript プロジェクトでは **、./src/taskpane/taskpane.js** ファイルには、アドインを初期化するコードが含まれている。 TypeScript プロジェクトでは **、./src/taskpane/taskpane.ts** ファイルには、アドインを初期化するコードと、Office JavaScript API ライブラリを使用して Microsoft Graph から Office ドキュメントにデータを追加するコードが含まれている。

### <a name="authentication"></a>認証

次のファイルは、SSO プロセスを容易にし、ドキュメントにデータをOfficeします。

- JavaScript プロジェクトでは **、./src/helpers/documentHelper.js** ファイルには、Office JavaScript API ライブラリを使用して Microsoft Graph から Office ドキュメントにデータを追加するコードが含まれている。 TypeScript プロジェクトにはそのようなファイルはありません。Office JavaScript API ライブラリを使用して Microsoft Graph から Office ドキュメントにデータを追加するコードは、代わりに **./src/taskpane/taskpane.ts** に存在します。

- **./src/helpers/fallbackauthdialog.html** ファイルは、フォールバック認証戦略の JavaScript を読み込む UI レス ページです。

- **./src/helpers/fallbackauthdialog.js** ファイルには、ユーザーにサインインするフォールバック認証戦略用の JavaScript がmsal.js。

- **./src/helpers/fallbackauthhelper.jsファイル** には、SSO 認証がサポートされていないシナリオでフォールバック認証戦略を呼び出す作業ウィンドウ JavaScript が含まれています。

- **./src/helpers/ssoauthhelper.js** ファイルには、SSO API `getAccessToken` へのJavaScript 呼び出しが含まれ、ブートストラップ トークンの受信し、Microsoft Graph へのアクセス トークンのブートストラップ トークン交換の開始、データのための Microsoft Graph への呼び出しを行います。
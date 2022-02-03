### <a name="configuration"></a>構成

次のファイルは、アドインの構成設定を指定します。

- プロジェクトのルート ディレクトリにある **./manifest.xml** ファイルで、アドインの機能と設定を定義します。

- **./.ENV** ファイルはプロジェクトのルート ディレクトリにあり、アドイン プロジェクトで使用される定数を定義します。

### <a name="task-pane"></a>作業ウィンドウ

次のファイルは、アドインの作業ウィンドウ UI と機能を定義します。

- **./src/taskpane/taskpane.html** ファイルには、作業ペイン用のHTMLマークアップが含まれています。

- **./src/taskpane/taskpane.css** ファイルには、作業ペインのコンテンツに適用されるCSSが含まれています。

- JavaScript プロジェクトでは、**./src/taskpane/taskpane.js** ファイルにアドインを初期化するコードが含まれています。 TypeScript プロジェクトでは、**./src/taskpane/taskpane.ts** ファイルに、アドインを初期化するコードと、Office JavaScript API ライブラリを使用して Microsoft Graph からのデータを Office ドキュメントに追加するコードが含まれています。

### <a name="authentication"></a>認証

次のファイルは、SSO 処理のサポートを行い、Office ドキュメントにデータを書き込みます。

- JavaScript プロジェクトでは、**./src/helpers/documentHelper.js** ファイルに、Office JavaScript API ライブラリを使用して、Microsoft Graph から Office ドキュメントにデータを追加するコードが含まれています。 TypeScript プロジェクトにはそのようなファイルは存在せず、代わりに Office JavaScript API ライブラリを使用して Microsoft Graph からのデータを **./src/taskpane/taskpane.ts** に存在する Office ドキュメントに追加するコードが含まれます。

- **./src/helpers/fallbackauthdialog.html** ファイルは、フォールバック認証戦略の JavaScript をロードする UI のないページです。

- **./src/helpers/fallbackauthdialog.js** ファイルには、msal.js でユーザーにサインオンするフォールバック認証用の JavaScript が含まれています。

- **./src/helpers/fallbackauthhelper.js** ファイルには、SSO 認証がサポートされていないシナリオでフォールバック認証戦略を呼び出す作業ウィンドウの JavaScript が含まれています。

- **./src/helpers/ssoauthhelper.js** ファイルには、SSO API、`getAccessToken` への JavaScript 呼び出しが含まれ、アクセス トークンを受信し、Microsoft Graph へのアクセス許可を持つ新しいアクセス トークン用のアクセス トークン交換の開始、データ用の Microsoft Graph への呼び出しを行います。
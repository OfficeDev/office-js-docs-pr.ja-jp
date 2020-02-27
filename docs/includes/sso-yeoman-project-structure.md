### <a name="configuration"></a>構成

次のファイルは、アドインの構成設定を指定します。

- プロジェクトのルートディレクトリにある **./ manifest.xml**ファイルは、アドインの設定と機能性を定義します。

- **./.ENV**プロジェクトのルートディレクトリにあるファイルには、アドインプロジェクトで使用される定数が定義されています。

### <a name="task-pane"></a>作業ウィンドウ 

次のファイルは、アドインの作業ウィンドウの UI と機能を定義します。

- **./src/taskpane/taskpane.html**ファイルには、作業ペイン用のHTMLマークアップが含まれています。

- **./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。

- JavaScript プロジェクトでは、 **/src/taskpane/taskpane.js**ファイルにアドインを初期化するコードが含まれています。 TypeScript プロジェクトでは、/src/taskpane/taskpane.ts ファイルにアドインを初期化するコードと、Office JavaScript ライブラリを使用して Microsoft Graph から Office ドキュメントにデータを追加するコードも記述されてい**ます**。

### <a name="authentication"></a>認証

次のファイルにより、SSO プロセスが容易になり、Office ドキュメントにデータが書き込まれます。

- JavaScript プロジェクトの/Src/helpers/documentHelper.js ファイルには、Office JavaScript ライブラリを使用して Microsoft Graph のデータを Office ドキュメントに追加するコードが含まれてい**ます**。 このようなファイルは TypeScript プロジェクトには含まれていません。Office JavaScript ライブラリを使用して Microsoft Graph から Office ドキュメントにデータを追加するコードは、代わりに **/src/taskpane/taskpane.ts**にあります。

- **./Src/helpers/fallbackauthdialog.html**ファイルは、フォールバック認証戦略の JavaScript を読み込む UI レスページです。

- **./Src/helpers/fallbackauthdialog.js**ファイルには、msal .js を使用してユーザーにサインするフォールバック認証戦略の JavaScript が含まれています。

- **/Src/helpers/fallbackauthhelper.js**ファイルには、SSO 認証がサポートされていないシナリオでフォールバック認証戦略を呼び出す作業ウィンドウ JavaScript が含まれています。

- **./src/helpers/ssoauthhelper.js** ファイルには、SSO API `getAccessToken` へのJavaScript 呼び出しが含まれ、ブートストラップ トークンの受信し、Microsoft Graph へのアクセス トークンのブートストラップ トークン交換の開始、データのための Microsoft Graph への呼び出しを行います。
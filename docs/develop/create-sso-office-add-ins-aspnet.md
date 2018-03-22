---
title: シングル サインオンを使用する ASP.NET Office アドインを作成する
description: null
ms.date: 01/23/2018
---

# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>シングル サインオンを使用する ASP.NET Office アドインを作成する (プレビュー)

ユーザーが Office にサインインしたとき、アドインは同じ資格情報を使用し、再度のサインインを要求することなく、複数のアプリケーションへのアクセスを許可することができます。概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。

この記事では、.NET 対応の ASP.NET、OWIN、および Microsoft 認証ライブラリ (MSAL) を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。

> [!NOTE]
> Node.js ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

* 入手可能な Visual Studio 2017 プレビューの最新バージョン。

* Office 2016 バージョン 1708、ビルド 8424.nnnn 以降 (「クイック実行」と呼ばれることもある Office 365 のサブスクリプション バージョン)。このバージョンを入手するには、Office Insider への参加が必要になることがあります。詳細については、「[Office Insider](https://products.office.com/ja-jp/office-insider?tab=tab-1)」を参照してください。

## <a name="set-up-the-starter-project"></a>スタート プロジェクトをセットアップする

1. 「[Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso)」にあるリポジトリを複製するかダウンロードします。

1. **[Before]** フォルダーを開いて、Visual Studio で .sln ファイルを開きます。これがスタート プロジェクトになります。SSO や承認に直接関連しない UI などの側面は、既に完了しています。

    > [!NOTE]
    > 同じリポジトリ内には、サンプルの完成版も含まれています。これは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。完成版を使用する場合は、`sln` ファイルを開いて、この記事の手順をそのまま実行しますが、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションは省略してください。

1. プロジェクトを開いたら、そのプロジェクトを Visual Studio でビルドします。その結果として、packages.config ファイルにリストされたパッケージがインストールされます。コンピューターのローカル パッケージ キャッシュに含まれるパッケージの数に応じて、数秒から数分の時間がかかります。

    > [!NOTE]
    > ID 名前空間に関するエラーが表示されます。 これは構成の問題の副作用ですが、次のステップで修正します。 重要な点は、パッケージがインストールされていることです。

1. 現在、SSO (バージョン `1.1.1-alpha0393`) に必要な MSAL ライブラリ (Microsoft.Identity.Client) は標準の nuget カタログの一部ではないため、package.config にはリストされていません。これは、個別にインストールする必要があります。 

   > 1. **[ツール]** メニューで **[Nuget パッケージ マネージャー]** > **[パッケージ マネージャー コンソール]** に移動します。 

   > 2. コンソールで、次のコマンドを実行します。 これは高速インターネット接続の場合でも、完了までに数分かかることがあります。 完了すると、コンソールの出力の末尾に **'Microsoft.Identity.Client 1.1.1-alpha0393' が正常にインストールされました...** というメッセージが表示されます。

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.1-alpha0393 -Source https://www.myget.org/F/aad-clients-nightly/api/v3/index.json`

   > 3. **ソリューション エクスプローラー**で **[参照]** を右クリックします。**Microsoft.Identity.Client** がリストされていることを確認します。リストされていない場合やエントリに警告アイコンが表示されている場合は、エントリを削除してから Visual Studio 参照の追加ウィザードを使用して、**... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.1-alpha0393\lib\net45\Microsoft.Identity.Client.dll** のアセンブリへの参照を追加します。

1. もう一度プロジェクトをビルドします。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Azure AD v2.0 エンドポイントにアドインを登録する

1. [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com) に移動します。

1. 管理者の資格情報を使用して Office 365 テナントにサインインします。たとえば、MyName@contoso.onmicrosoft.com

1. **[アプリの追加]** をクリックします。

1. ダイアログが表示されたら、アプリ名として「Office-Add-in-ASPNET-SSO」を使用して、**[アプリケーションの作成]** をクリックします。

1. アプリの構成ページが開いたら、**[アプリケーション ID]** をコピーして保存します。これは、この後の手順で使用します。

    > [!NOTE]
    > この ID は、Office ホスト アプリケーション (たとえば、PowerPoint、Word、Excel) などの別のアプリケーションが、このアプリケーションへの承認されたアクセスを求めるときの「対象ユーザー」値になります。また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。

1. **[アプリケーション シークレット]** セクションで、**[新しいパスワードを生成する]** をクリックします。新しいパスワード (「アプリケーション シークレット」とも呼びます) が示されたポップアップ ダイアログが開きます。*このパスワードをすぐにコピーして、アプリケーション ID と共に保存します。*これは、この後の手順で必要になります。その後で、ダイアログを閉じます。

1. **[プラットフォーム]** セクションで、**[プラットフォームの追加]** をクリックします。

1. 開いたダイアログで、**[Web API]** を選択します。

1. **[アプリケーション ID URI]** が、"api://{App ID GUID}" という形式で生成されています。二重スラッシュと GUID の間に、文字列 "localhost:44355/" を挿入します。この ID の全体は、`api://localhost:44355/{App ID GUID}` のようになります。 

    > [!NOTE]
    > **[アプリケーション ID URI]** の直後の **[スコープ]** 名のドメイン部分は、一致するように自動的に変更されます。 これは `api://localhost:44355/{App ID GUID}/access_as_user` のようになります。

1. **[事前承認済みアプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。 次のそれぞれの ID を事前承認する必要があります。 1 つの ID を入力するたびに、新しい空のテキスト ボックスが表示されます。 (GUID のみを入力してください。)
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. それぞれの **[アプリケーション ID]** の横の **[スコープ]** ドロップダウンを開いて、`api://localhost:44355/{App ID GUID}/access_as_user` のボックスをオンにします。

1. **[プラットフォーム]** セクションの上部にある **[プラットフォームの追加]** を再度クリックして、**[Web]** を選択します。

1. **[プラットフォーム]** の下側の新しい **[Web]** セクションで、**[リダイレクト URL]** として `https://localhost:44355` を入力します。

    > [!NOTE]
    > この記事の執筆時には、**[プラットフォーム]** セクションに **[Web API]** が表示されないことがあります。特に、**Web** プラットフォームを追加して*登録ページを保存*した後でページが最新の情報に更新されたときに発生します。**[Web API]** プラットフォームが登録に含まれていることを再確認する場合は、ページの下側にある **[アプリケーション マニフェストの編集]** ボタンをクリックします。マニフェストの **identifierUris** プロパティに、文字列 `api://localhost:44355/{App ID GUID}` が表示されている必要があります。また、値 `access_as_user` を保持している **value** サブプロパティがある **oauth2Permissions** プロパティもあります。

1. **[Microsoft Graph のアクセス許可]** セクションを下にスクロールして、**[委任されたアクセス許可]** サブセクションを表示します。**[追加]** ボタンを使用して、**[アクセス許可の選択]** ダイアログを開きます。

1. ダイアログ ボックスで、次の各アクセス許可のボックスをオンにします。 実際にアドイン自体に必要なのは最初のものだけですが、サーバー側コードで使用される MSAL ライブラリで `offline_access` および `openid` が必要とされます。 Office ホストがアドインの Web アプリケーションに対してトークンを取得するために、`profile` のアクセス許可が必要です。
    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > `User.Read` アクセス許可は既定でリストされている場合があります。必要でないアクセス許可は依頼しない方がよいため、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。

1. ダイアログの下部にある **[OK]** をクリックします。

1. 登録ページの下部にある **[保存]** をクリックします。

## <a name="grant-admin-consent-to-the-add-in"></a>アドインに管理者の同意を付与する

> [!NOTE]
> この手順が必要とされるのは、アドインを開発しているときだけです。実際に運用するアドインを AppSource またはアドイン カタログに展開した場合、インストール時に、各ユーザーが個別にそのアドインを信頼するか、管理者が組織のために同意することになります。

1. Visual Studio でアドインを実行していない場合は、**F5** キーを押してアドインを実行します。この手順をスムーズに完了するには、IIS で実行する必要があります。

1. 次に示す文字列内のプレースホルダー "{application_ID}" は、アドインの登録時にコピーしたアプリケーション ID に置き換えます: `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. そうしてできた URL をブラウザーのアドレス バーに貼り付けて、そこに移動します。

1. ダイアログが表示されたら、管理者の資格情報を使用して Office 365 テナントにサインインします。

1. その後で、Microsoft Graph データにアクセスするためのアクセス許可をアドインに付与するように求めるダイアログが表示されます。**[承諾]** をクリックします。

1. ブラウザー ウィンドウ (タブ) は、アドインの登録時に指定した **[リダイレクト URL]** にリダイレクトされ、アドインのホーム ページがブラウザーで開かれます。

2. ブラウザーのアドレスバーには、GUID 値の付いた "tenant" クエリ パラメーターが表示されます。これは、Office 365 テナントの ID です。この値をコピーして保存します。これは後の手順で使用します。

3. ウィンドウ (タブ) を閉じます。

1. Visual Studio のデバッガーを停止します。

## <a name="configure-the-add-in"></a>アドインを構成する

1. 次に示す文字列内のプレースホルダー "{tenant_ID}" は、前の手順で取得した Office 365 テナント ID に置き換えます。何らかの理由で、まだ ID を取得していない場合は、「[Office 365 テナント ID を検索する](https://support.office.com/ja-jp/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b)」に示したいずれかの方法を使用して ID を取得します。

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. Visual Studio で、web.config を開きます。**[appSettings]** セクションには、値を割り当てる必要のあるいくつかのキーがあります。

3. "ida:Issuer" という名前のキーの値として、手順 1 で作成した文字列を使用します。この値に、空白スペースが含まれていないことを確認してください。

4. 次に示す値を対応するキーに代入します。

    |キー|値|
    |:-----|:-----|
    |ida:ClientID|アドインの登録時に取得したアプリケーション ID。|
    |ida:Audience|アドインの登録時に取得したアプリケーション ID。|
    |ida:Password|アドインの登録時に取得したパスワード。|

   次に、4 つのキーの変更後の例を示します。*ClientID と Audience が同じになっている点に注目してください*。両方の目的に単一のキーを使用することもできますが、これらは必ずしも同じではないため、別々に保持しておくと web.config のマークアップが再利用しやすくなります。また、別のキーを使用することで、アドインが Office ホストに関連する OAuth リソースと、Microsoft Graph に関連する OAuth クライアントの両方でであるという考えが補強されます。

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    
    ```

   > [!NOTE]
   > その他の **[appSettings]** セクションの設定は、未変更のままにします。

1. ファイルを保存して閉じます。

1. アドイン プロジェクトで、アドイン マニフェスト ファイル "Office-Add-in-ASPNET-SSO.xml" を開きます。

1. ファイルの最後までスクロールします。

1. `</VersionOverrides>` 終了タグの直前に、次に示すマークアップがあります。

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. このマークアップ内の*両方の場所の*プレースホルダー "{application_GUID here}" を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。「{」と「}」は ID の一部ではないため、これらを含めないでください。これは、web.config の ClientID と Audience に使用したものと同じ ID です。

    > [!NOTE]
    > * **[リソース]** の値は、アドインの登録に Web API プラットフォームを追加したときに設定した **[アプリケーション ID URI]** です。
    > * **[範囲]** セクションは、アドインが AppSource から販売された場合に、同意ダイアログ ボックスを生成するためにのみ使用します。

1. Visual Studio で、**[エラー一覧]** の **[警告]** タブを開きます。 `<WebApplicationInfo>` が `<VersionOverrides>` の有効な子ではないという警告が表示される場合は、Visual Studio 2017 プレビューのバージョンで SSO マークアップが認識されていません。 回避策として、Word、Excel、または PowerPoint のアドインに対して、次の操作を行います。 (Outlook アドインを使用している場合は、以下の回避策を参照してください。)

   - **Word、Excel、および PowerPoint の回避策**

        1. マニフェストの `</VersionOverrides>` の終了タグの直前の `<WebApplicationInfo>` セクションをコメント アウトします。

        2. F5 キーを押してデバッグ セッションを開始します。これにより、次のフォルダーにマニフェストのコピーが作成されます (これには、Visual Studio よりも**ファイル エクスプローラー**の方が容易にアクセスできます): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

        3. マニフェストのコピーから、`<WebApplicationInfo>` セクションの周囲のコメント構文を削除します。

        4. マニフェストのコピーを保存します。

        5. この時点で、次回 F5 キーを押したときに、このマニフェストのコピーが Visual Studio によって上書きされないようにする必要があります。**ソリューション エクスプローラー**の上部にあるソリューション ノード (どちらのプロジェクト ノードでもない) を右クリックします。

        6. コンテキスト メニューから **[プロパティ]** を選択します。**[ソリューション プロパティ ページ]** ダイアログ ボックスが開きます。

        7. **[構成プロパティ]** を展開し、**[構成]** を選択します。

        8. **Office-Add-in-ASPNET-SSO** プロジェクト (**Office-Add-in-ASPNET-SSO-WebAPI** プロジェクトでは*ありません*) の行で、**[ビルド]** と **[展開]** を選択解除します。

        9. **[OK]** をクリックしてダイアログ ボックスを閉じます。

   - **Outlook の回避策**

        1. 開発用コンピューターで、既存の `MailAppVersionOverridesV1_1.xsd` を探します。 `./Xml/Schemas/{lcid}` の下の Visual Studio インストール ディレクトリに配置されています。 たとえば、英語版 (米国) の VS 2017 32 ビットの標準インストールの場合、完全なパスは、`C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033` になります。

        2. 既存のファイルの名前を、`MailAppVersionOverridesV1_1.old` に変更します。

        3. 変更したこのファイルを、フォルダーにコピーします。[変更済みの MailAppVersionOverrides スキーマ](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. Visual Studio でメインのマニフェスト ファイルを保存して閉じます。

## <a name="code-the-client-side"></a>クライアント側のコードの作成

1. **[Scripts]** フォルダー内の Home.js ファイルを開きます。これには、一部のコードが既に含まれています。
    * `Office.initialize` メソッドへの割り当てが、`getGraphAccessTokenButton` ボタン クリック イベントへのハンドラーの割り当てになります。
    * `showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。
    * `logErrors` メソッドは、エンド ユーザーを対象としていないエラーをコンソールにログ出力するものです。

1. `Office.initialize` への割り当ての下に、次に示すコードを追加します。このコードについては、次の点に注意してください。

    * アドインのエラー処理により、アクセス トークンの取得が別のオプションのセットを使用して自動的に再試行されることがあります。 カウンター変数 `timesGetOneDriveFilesHasRun` とフラグ変数 `triedWithoutForceConsent` を使用して、失敗するトークン取得の繰り返しからユーザーが抜け出せるようにします。 
    * この後の手順では `getDataWithToken` メソッドを作成しますが、そのメソッドで `forceConsent` というオプションが `false` に設定される点に注意してください。詳細については、次の手順で説明します。

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. `getOneDriveFiles` メソッドの下に、次のコードを追加します。このコードについては、次の点に注意してください。

    * `getAccessTokenAsync` は Office.js の新しい API です。これにより、アドインは Office ホスト アプリケーション (Excel、PowerPoint、Word など) に、アドインへのアクセス トークン (Office にサインインしているユーザーのトークン) を要求できるようになります。その Office ホスト アプリケーションが、Azure AD 2.0 エンドポイントにトークンを要求します。アドインの登録時に、アドインに対する Office ホストを事前認証しているため、Azure AD はトークンを送信します。
    * Office にサインインしているユーザーがいない場合、Office ホストはユーザーにサインインを求めるダイアログを表示します。
    * オプションのパラメーター `forceConsent` を `false` に設定すると、ユーザーがアドインを使用するたびに、Office ホストにアドインへのアクセス権を付与するための同意を求めるダイアログが表示されなくなります。 ユーザーが初めてアドインを実行すると、`getAccessTokenAsync` の呼び出しは失敗しますが、この後の手順で追加するエラー処理ロジックにより、`forceConsent` オプションを `true` に設定した再呼び出しが自動的に実行され、ユーザーに同意を求めるダイアログが表示されます。ただし、これは初回時のみ実行されます。
    * `handleClientSideErrors` メソッドは、この後の手順で作成します。

    ```javascript
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. TODO1 を次に示す行に置き換えます。`getData` メソッドとサーバー側の "/api/values" ルートは、この後の手順で作成します。エンドポイントには、相対 URL を使用します。これは、その URL がアドインと同じドメインでホストされている必要があるためです。

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. `getOneDriveFiles` メソッドの下に、以下を追加します。このコードについては、次の点に注意してください。

    * このメソッドは、特定の Web API エンドポイントを呼び出して、Office ホスト アプリケーションがアドインへのアクセスに使用したものと同じアクセス トークンを渡します。サーバー側では、このアクセス トークンが Microsoft Graph へのアクセス トークンを取得するための「代理 (on-behalf-of)」フローで使用されます。
    * `handleServerSideErrors` メソッドは、この後の手順で作成します。

    ```javascript
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        }); 
    }
    ```

### <a name="create-the-error-handling-methods"></a>エラー処理のメソッドを作成する

1. `getData` メソッドの下に、次のメソッドを追加します。 このメソッドは、Office ホストがアドインの Web サービスへのアクセス トークンを取得できないときに、アドインのクライアントでエラーを処理します。 こうしたエラーはエラー コードで報告されるため、このメソッドでは `switch` ステートメントを使用してエラーを識別します。

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. `TODO2` を次のコードに置き換えます。 エラー 13001 は、ユーザーがログインしていない場合、または 2 番目の認証要素の指定を求めるダイアログに応答しないでキャンセルした場合に発生します。 どちらの場合も、このコードでは `getDataWithToken` メソッドを再実行して、サインインを求めるダイアログの表示を強制するようにオプションを設定します。

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. `TODO3` を次のコードに置き換えます。 エラー 13002 は、ユーザーのサインインまたは同意が中断された場合に発生します。 ユーザーに対して 1 回だけ再試行を求めます。

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. `TODO4` を次のコードに置き換えます。 エラー 13003 は、ユーザーが職場または学校アカウントと、Micrososoft アカウントのどちらでもないアカウントでログインしている場合に発生します。 ユーザーに対して、サインアウトしてからサポートされているアカウントの種類で再度サインインするように求めます。

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > エラー 13004 と 13005 は、開発時にのみ発生するため、このメソッドでは処理しません。 これらは、ランタイム コードで修正できるものではなく、エンド ユーザーに報告しても意味がありません。

1. `TODO5` を次のコードと置き換えます。エラー 13006 は、Office ホストで未指定のエラーがある場合に発生します。ホストが不安定な状態にあることを示している可能性があります。ユーザーに Office の再起動を求めます。

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. `TODO6` を次のコードに置き換えます。 エラー 13007 は、Office ホストの AAD との相互作用に問題があり、ホストがアドイン Web サービス/アプリケーションへのアクセス トークンを取得できない場合に発生します。 ネットワークに一時的な問題が発生している可能性があります。 しばらく待ってから再試行するようにユーザーに求めます。

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. `TODO7` を次のコードと置き換えます。エラー 13008 は、前回の `getAccessTokenAsync` の呼び出しが完了する前に、それを呼び出す操作をユーザーがトリガーしたときに発生します。

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. `TODO8` を次のコードに置き換えます。 エラー 13009 は、アドインが強制的な同意をサポートしていないときに、`forceConsent` オプションを `true` に設定して `getAccessTokenAsync` を呼び出した場合に発生します。 通常、この場合は、コードによって同意オプションを `false` に設定して自動的に `getAccessTokenAsync` を再実行する必要があります。 ただし、`forceConsent` を `true` に設定してメソッドを呼び出すこと自体が、そのオプションを `false` に設定したメソッドの呼び出しで発生したエラーに対する自動的な応答の場合もあります。 その場合は、コードで再試行するのではなく、ユーザーにサインアウトしてから再度サインインするように通知する必要があります。

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. `TODO9` を次のコードに置き換えます。

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. `handleClientSideErrors` メソッドの下に、次のメソッドを追加します。このメソッドは、代理 (on-behalf-of) フローの実行時または Microsoft Graph からのデータの取得時の問題により、アドインの Web サービスで発生したエラーを処理します。

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Parse the JSON response.

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO14: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. `TODO10` を次のコードに置き換えます。 アドインの Web サービスがアドインのクライアント側に渡すほとんどの `4xx` エラーには、その応答内に **ExceptionMessage** プロパティが含まれています。このプロパティには、AADSTS (Azure Active Directory Secure Token Service) エラー番号などのデータが格納されています。 ただし、AAD がアドインの Web サービスに追加の認証要素を求めるメッセージを送信するときには、そのメッセージに特殊な **Claims** プロパティが含まれます。このプロパティによって、どの追加要素が必要になるかが (コード番号で) 示されます。 HTTP 応答を作成してクライアントに送信する ASP.NET API は、この **Claims** プロパティを認識しないため、このプロパティを応答オブジェクトに含めません。 この後の手順で作成するサーバー側のコードでは、これに対処するために、手動で応答オブジェクトに **Claims** 値を追加しています。 この値は、**Message** プロパティに含めるため、コードでは、そのプロパティも解析する必要があります。

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    }
    ```

1. `TODO11` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * エラー 50076 は、Microsoft Graph が認証の追加フォームを必要とする場合に発生します。
    * Office ホストは、`authChallenge` オプションとして **Claims** 値を使用して新しいトークンを取得します。 これにより、認証のすべての必要なフォームをユーザーに表示するように AAD に指示します。 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. `TODO12` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * エラー 65001 は、1 つ以上のアクセス許可について Microsoft Graph にアクセスするための同意が与えられていない (または取り消されている) ことを意味します。 
    * アドインでは、`forceConsent` オプションを `true` に設定して新しいトークンを取得する必要があります。

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
       // getDataWithToken({ forceConsent: true });
    }    
    ```

1. `TODO13` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * エラー 70011 には複数の意味があります。無効なスコープ (アクセス許可) が要求されていることを意味する場合、このアドインに重要となります。コードでは番号だけでなくエラーの説明全体を確認します。
    * アドインでは、エラーを報告する必要があります。

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. `TODO14` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * この後の手順で作成するサーバー側のコードでは、アドインのクライアントが AAD に送信して代理 (on-behalf-of) フローで使用されるアクセス トークンに `access_as_user` スコープ (アクセス許可) が含まれていない場合に、メッセージ `Missing access_as_user` を送信します。
    * アドインでは、エラーを報告する必要があります。

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. `TODO15` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * サーバー側のコードで使用する ID ライブラリ (Microsoft Authentication Library - MSAL) では、期限切れのトークンや無効なトークンが Microsoft Graph に送信されないようにする必要があります。ただし、その事態が発生した場合は、アドインの Web サービスに Microsoft Graph から返されるエラーにコード `InvalidAuthenticationToken` が含まれています。後の手順で作成するサーバー側のコードは、このメッセージをアドインのクライアントに中継します。
    * この場合、アドインはカウンター変数とフラグ変数をリセットしてから、ボタン ハンドラー メソッドを再呼び出しすることで、認証プロセス全体を最初から開始する必要があります。

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. `TODO16` を次のコードに置き換えます。

    ```javascript
    else {
        logError(result);
    }    
    ```

1. ファイルを保存して閉じます。

## <a name="code-the-server-side"></a>サーバー側のコードを作成する

### <a name="configure-the-owin-middleware"></a>OWIN ミドルウェアを構成する

1. プロジェクトのルートにある Startup.cs を開きます。

1. Startup クラスの宣言にキーワード `partial` を追加します (まだ追加されていない場合)。これは、次のようになります。

    `public partial class Startup`

1. `Configuration` メソッドの本文に、次に示す行を追加します。`ConfigureAuth` メソッドは、この後の手順で作成します。

    `ConfigureAuth(app);`

1. ファイルを保存して閉じます。

1. **App_Start** フォルダーを右クリックして、**[追加] > [クラス]** を選択します。

1. **[新しい項目の追加]** ダイアログで、ファイルに「**Startup.Auth.cs**」という名前を付けて **[追加]** をクリックします。

1. 新しいファイルで名前空間の名前を `Office_Add_in_ASPNET_SSO_WebAPI` に短縮します。

1. ファイルの先頭に、次に示す `using` ステートメントがすべて揃っていることを確認します。

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. `Startup` クラスの宣言にキーワード `partial` を追加します (まだ追加されていない場合)。これは、次のようになります。

    `public partial class Startup`

1. 次に示すメソッドを `Startup` クラスに追加します。このメソッドでは、クライアント側の Home.js ファイルの `getData` メソッドから渡されたアクセス トークンを OWIN ミドルウェアで検証する方法を指定します。承認プロセスは、`[Authorize]` 属性で修飾された Web API エンドポイントが呼び出されたときには必ずトリガーされます。

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. TODO3 を次のように置き換えます。 このコードの注意点は次のとおりです。

    * このコードでは、Office ホストから得られるアクセス トークン (`getData` のクライアント側呼び出しによって渡されるトークン) で指定された対象ユーザーとトークン発行者が web.config で指定された値と一致する必要があることを OWIN に指示します。
    * `SaveSigninToken` を `true` に設定することで、OWIN は Office ホストからの Raw トークンを保存するようになります。これは、アドインが「代理」フローで Microsoft Graph へのアクセス トークンを取得するために必要になります。
    * OWIN ミドルウェアでは、スコープは検証されません。`access_as_user` が含まれている必要があるアクセス トークンのスコープは、コントローラーで検証されます。

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. TODO4 を次のように置き換えます。このコードの注意点は次のとおりです。

    * より一般的な `UseWindowsAzureActiveDirectoryBearerAuthentication` は Azure AD V2 エンドポイントに準拠していないため、その代わりとしてメソッド `UseOAuthBearerAuthentication` が呼び出されます。
    * このメソッドに渡される探索 URL は、Office ホストから受け取ったアクセス トークンの署名の検証に必要になるキーを取得するための方法を OWIN ミドルウェアが取得する場所になります。

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. ファイルを保存して閉じます。

### <a name="create-the-apivalues-controller"></a>/api/values コントローラーを作成する

1. ファイル **Controllers\ValueController.cs** を開きます。

2. ファイルの先頭に、次に示す `using` ステートメントがあることを確認します。

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. `ValuesController` を宣言している行のすぐ上に、属性 `[Authorize]` を追加します。これにより、アドインはコントローラー メソッドが呼び出されたときに、最後の手順で構成した承認プロセスを必ず実行するようになります。アドインへの有効なアクセス トークンを持つ呼び出し元のみが、コントローラーのメソッドを起動できます。

    > [!NOTE]
    > 運用環境の ASP.NET MVC Web API サービスには、1 つ以上のカスタム [FilterAttribute](https://msdn.microsoft.com/ja-jp/library/system.web.http.filters(v=vs.108).aspx) クラスに代理 (on-behalf-of) フロー用のカスタム ロジックを用意する必要があります。 この学習用サンプルでは、メイン コントローラーにロジックを配置して、認証とデータのフェッチ ロジックの全体的なフローを簡単に把握できるようにしています。 さらに、このサンプルが「[Azure Samples](https://github.com/Azure-Samples/)」の承認サンプルのパターンと一致するようになります。    

4. 次のメソッドを `ValuesController` に追加します。 戻り値は、`Task<IEnumerable<string>>` ではなく `GET api/values` メソッドでより一般的な `Task<HttpResponseMessage>` になる点に注意してください。 これは、カスタムの承認ロジックがコントローラー内にあることの副作用です。そのロジックの一部のエラー条件では、HTTP 応答オブジェクトをアドインのクライアントに送信することが必要になります。 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. `TODO1` を次のコードに置き換えます。このコードでは、`access_as_user` を含むトークンで指定されているスコープを検証します。

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > `access_as_user` スコープのみを使用して、Office アドインの代理 (on-behalf-of) フローを処理する API を承認する必要があります。サービス内の他の API は、独自のスコープ要件が必要です。これにより、Office が取得するトークンでアクセスできるものが制限されます。

6. `TODO2` を次のコードに置き換えます。このコードの注意点は次のとおりです。
    * このコードでは、Office ホストから受け取った Raw アクセス トークンを別のメソッドに渡される `UserAssertion` オブジェクトに変換します。
    * アドインは、Office ホストとユーザーがアクセスする必要のあるリソース (または対象ユーザー) の役割を果たさなくなります。この時点で、それ自体が Microsoft Graph にアクセスする必要があるクライアントになります。`ConfidentialClientApplication` は MSAL の「クライアント コンテキスト」オブジェクトになります。
    * `ConfidentialClientApplication` コンストラクターへの 3 番目のパラメーターはリダイレクト URL です。これは、実際には「代理」フローで使用されることはありませんが、正しい URL を使用することをお勧めします。4 番目と 5 番目のパラメーターは、永続ストアを定義するために使用できます。このストアにより、有効期限が切れていないトークンをアドインの異なるセッション間で再使用できるようになります。このサンプルでは、永続ストアは実装していません。
    * MSAL では `openid`、`offline_access` の各スコープが機能することが必要ですが、コードがこれらを重複して要求するとエラーがスローされます。 コードが `profile` を要求した場合にもエラーがスローされます。それは、実際には Office ホスト アプリケーションがアドインの Web アプリケーションに対しトークンを取得するときだけに使用します。 そのため、`Files.Read.All` のみが明示的に要求されます。

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. `TODO3` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` メソッドは、最初にメモリ内の MSAL キャッシュで一致するアクセス トークンを探します。それが見つからなかった場合にのみ、Azure AD V2 エンドポイントとの「代理」フローを開始します。
    * MS Graph リソースが多要素認証を必要とし、ユーザーがまだそれを提供していない場合、AAD は Claims プロパティが含まれている例外をスローします。
    * Claims プロパティの値は、クライアントに渡す必要があります。クライアントは、その値を Office ホストに渡します。Office ホストは、その値を新しいトークンの要求に含めます。AAD は、認証のすべての必要なフォームをユーザーに示します。
    * `MsalServiceException` 以外の種類の例外は、意図的にキャッチしていないため、`500 Server Error` メッセージとしてクライアントに伝達されます。

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {        
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

8. `TODO3a` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * MS Graph リソースが多要素認証を必要としているときに、その認証をユーザーがまだ指定していない場合、AAD はエラー AADSTS50076 と **Claims** プロパティを含む "400 Bad Request" を返します。MSAL は **MsalUiRequiredException** (**MsalServiceException** から継承) をこの情報とともにスローします。 
    * **Claims** プロパティの値は、クライアントに渡す必要があります。クライアントは、その値を Office ホストに渡します。Office ホストは、その値を新しいトークンの要求に含めます。AAD は、認証のすべての必要なフォームのための指示をユーザーに示します。
    * 例外から HTTP 応答を作成する API は、**Claims** プロパティを認識しないため、このプロパティを応答オブジェクトに含めません。 これが含まれたメッセージを手動で作成する必要があります。 ただし、カスタムの **Message** プロパティは **ExceptionMessage** プロパティの作成を妨げるため、クライアントがエラー ID `AADSTS50076` を取得するには、その ID をカスタムの **Message** に追加する以外に方法はありません。 クライアントの JavaScript では、応答に **Message** または **ExceptionMessage** が含まれているかどうかを検出する必要があるため、どちらを読み取るかを認識します。
    * カスタム メッセージは、JSON として書式設定されているため、クライアント側の JavaScript は既知の `JSON` オブジェクトのメソッドでメッセージを解析できます。
    * `SendErrorToClient` メソッドは、この後の手順で作成します。 2 番目のパラメーターは、**Exception** オブジェクトです。 この場合、コードは `null` を渡します。これは、**Exception** オブジェクトが含まれていることで、生成される HTTP 応答には **Message** プロパティが含められなくなるためです。

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. `TODO3b` と `TODO3c` を次のコードに置き換えます。このコードの注意点は次のとおりです。

    * AAD の呼び出しにユーザーまたはテナント管理者のどちらも同意していない (または同意が取り消された) スコープ (アクセス許可) が少なくとも 1 つ含まれていると、 AAD はエラー `AADSTS65001` と共に "400 Bad Request" を返します。 MSAL は、この情報と共に **MsalUiRequiredException** をスローします。 クライアントは、オプション `{ forceConsent: true }` を使用して `getAccessTokenAsync` を再呼び出しする必要があります。
    *  AAD の呼び出しに AAD が認識しないスコープが少なくとも 1 つ含まれていると、AAD はエラー `AADSTS70011` と共に "400 Bad Request" を返します。 MSAL は、この情報と共に **MsalUiRequiredException** をスローします。 クライアントは、ユーザーに通知する必要があります。
    *  すべての説明が含まれている理由は、別の条件で 70011 が返されたときに、このアドインでは無効なスコープの存在を意味する場合のみを処理する必要があるためです。 
    *  **MsalUiRequiredException** オブジェクトが `SendErrorToClient` に渡されます。これにより、エラー情報を格納している **ExceptionMessage** プロパティが HTTP 応答に含まれるようにします。
    *  カスタム メッセージは存在しないため、3 番目のパラメーターでは `null` が渡されます。

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. `TODO3d` を次のコードに置き換えます。 このコードでは、**HttpStatusCode.Forbidden** (401) によるカスタムの HTTP 応答で例外を中継するのではなく、例外を再スローしています。 これにより、ASP.NET はステータス "500 Server Error" による独自の HTTP 応答を送信するようになります。

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. `TODO4` を次のように置き換えます。このコードの注意点は次のとおりです。

    * `GraphApiHelper` クラスと `ODataHelper` クラスは、**[Helpers]** フォルダー内のファイルで定義されています。`OneDriveItem` クラスは、**[Models]** フォルダー内のファイルで定義されています。これらのクラスについての詳しい説明は、承認や SSO に関連していないため、この記事の対象外になります。
    * 実際に必要なデータのみを Microsoft Graph に要求することでパフォーマンスが向上します。そのため、このコードでは、` $select` クエリ パラメーターで name プロパティのみが必要なことを指定し、`$top` パラメーターで最初の 3 つのフォルダー名またはファイル名のみが必要なことを指定しています。
    * Microsoft Graph に送信したトークンが無効な場合、Microsoft Graph は、コード "InvalidAuthenticationToken" を含む "401 Unauthorized" エラーを送信します。 その後で、ASP.NET は **RuntimeBinderException** をスローします。 これは、トークンの有効期限が切れているときにも発生しますが、MSAL では、そのような事態にならないようにする必要があります。 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);                    
    }
    ```

12. `TODO5` を次のように置き換えます。このコードの注意点は次のとおりです。 

    * 上記のコードでは OneDrive アイテムの *name* プロパティのみを要求していますが、Microsoft Graph は常に OneDrive アイテムの *eTag* プロパティを含めます。クライアントに送信するペイロードを縮小するために、次に示すコードではアイテム名のみで結果を再構築しています。
    * 3 つの OneDrive ファイルとフォルダーのリストは、"200 OK" HTTP 応答としてクライアントに送信されます。

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames); 
    return response;
    ```

13. Get メソッドの下に、次のメソッドを追加します。 このコードの注意点は次のとおりです。  

    * このメソッドは、サーバー側の例外に関する情報をクライアントに中継します。 
    * このメソッドに元の例外が渡されると、HttpError コンストラクターは例外オブジェクトからの情報を **ExceptionMessage** プロパティに含めます。  
    * 例外として `null` が渡されると、HttpError コンストラクターはメッセージ パラメーターを **Message** プロパティに含めます。**ExceptionMessage** プロパティは存在しなくなります。

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }        
    ```

## <a name="run-the-add-in"></a>アドインを実行する

1. 結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。

1. Visual Studio で、F5 キーを押します。PowerPoint が開き、**[ホーム]** リボンに **[SSO ASP.NET]** グループが表示されます。

1. このグループ内の **[アドインの表示]** ボタンをクリックすると、作業ウィンドウにアドインの UI が表示されます。

1. **[OneDrive からファイルを取得]** ボタンをクリックします。Office にサインインしていない場合は、サインインを求めるダイアログが表示されます。
    
    > [!NOTE]
    > 以前に別の ID で Office にサインオンしていて、そのときに開いたいくつかの Office アプリケーションが引き続き開いている場合、Office がその ID を確実に変更するとは限りません (PowerPoint で ID が変更済みのように表示されている場合でも)。 このような場合は、Microsoft Graph への呼び出しが失敗するか、以前の ID からのデータが返される可能性があります。 これを防止するには、必ず*他のすべての Office アプリケーションを閉じて*から、**[OneDrive からファイルを取得]** を押します。

1. サインインすると、ボタンの下に OneDrive のファイルとフォルダーのリストが表示されます。これには、15 秒以上かかることがあります (特に初回実行時)。

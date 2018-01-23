# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>シングル サインオンを使用する ASP.NET Office アドインを作成する

ユーザーが Office にサインインしたとき、アドインは同じ資格情報を使用し、再度のサインインを要求することなく、複数のアプリケーションへのアクセスを許可することができます。 概要については、「[Office アドインで SSO を有効化する](../develop/sso-in-office-add-ins.md)」を参照してください。

この記事では、.NET 対応の ASP.NET、OWIN、および Microsoft 認証ライブラリ (MSAL) を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。

> **注:**Node.js ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する Node.js Office アドインの作成](../develop/create-sso-office-add-ins-nodejs.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

* 入手可能な Visual Studio 2017 プレビューの最新バージョン。

>**メモ:**Visual Studio 2017 プレビューの最新バージョンは、現在、SSO に必要なアドイン マニフェスト マークアップと互換性がありません。 これを回避する方法の詳細は、次の手順で説明します。

* Office 2016 バージョン 1708、ビルド 8424.nnnn 以降 (「クイック実行」と呼ばれることもある Office 365 のサブスクリプション バージョン)。このバージョンを入手するには、Office Insider への参加が必要になることがあります。詳細については、「[Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1)」を参照してください。

## <a name="set-up-the-starter-project"></a>スタート プロジェクトをセットアップする

1. 「[Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso)」にあるリポジトリを複製するかダウンロードします。

1. **[Before]** フォルダーを開いて、Visual Studio で .sln ファイルを開きます。これがスタート プロジェクトになります。SSO や承認に直接関連しない UI などの側面は、既に完了しています。

    > メモ:同じリポジトリ内には、サンプルの完成版も含まれています。 これは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。 完成版を使用する場合は、*.sln ファイルを開いて、この記事の手順をそのまま実行しますが、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションは省略してください。

1. プロジェクトを開き、そのプロジェクトを Visual Studio でビルドすると、packages.config ファイルに列挙されたパッケージがインストールされます。 コンピューターのローカル パッケージ キャッシュに含まれるパッケージの数に応じて、数秒から数分程度の時間がかかります。

    > **重要!** Web API プロジェクトのルートにある packages.config は、Microsoft.Identity.Client (MSAL ライブラリ) のバージョン `1.1.1-alpha0393` を指定します。 最初に F5 キーを押した後に、このバージョンか、以降のバージョンがインストールされていることを確認する必要があります。**[ツール]** メニューから、**[Nuget パッケージ マネージャー]** > **[ソリューションの Nuget パッケージの管理...]** > **[インストール済みのパッケージ]** に移動します。 **Microsoft.Identity.Client** までスクロールし、インストールされたバージョンを確認します。 それが `1.1.1-alpha0393` より前のバージョンである (または **[インストール済みのパッケージ]** にない) 場合には、**[Nuget パッケージ マネージャー]** > **[パッケージ マネージャー コンソール]** に移動します。 コンソールで `Install-Package Microsoft.Identity.Client -Version 1.1.1-alpha0393 -Source https://www.myget.org/F/aad-clients-nightly/api/v3/index.json` コマンドを実行します。

1. プロジェクトのビルドが完了したら、F5 キーを押します。PowerPoint が開き、**[ホーム]** リボンに **[SSO ASP.NET]** グループが表示されます。

1. このグループ内の **[アドインの表示]** ボタンをクリックすると、作業ウィンドウにアドインの UI が表示されます。 この作業ウィンドウ内のボタンは、まだ機能に関連付けられていません。
2. Visual Studio で、デバッガーを停止します。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Azure AD v2.0 エンドポイントにアドインを登録する

1. [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com) に移動します。

1. 管理者の資格情報を使用して Office 365 テナントにサインインします。たとえば、MyName@contoso.onmicrosoft.com

1. **[アプリの追加]** をクリックします。

1. ダイアログが表示されたら、アプリ名として「Office-Add-in-ASPNET-SSO」を使用して、**[アプリケーションの作成]** をクリックします。

1. アプリの構成ページが開いたら、**[アプリケーション ID]** をコピーして保存します。 これは、この後の手順で使用します。

    > **メモ**:この ID は、Office ホスト アプリケーション (たとえば、PowerPoint、Word、Excel) などの別のアプリケーションが、このアプリケーションへの承認されたアクセスを求めるときの「対象ユーザー」値になります。 また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。

1. **[アプリケーション シークレット]** セクションで、**[新しいパスワードを生成]** をクリックします。 新しいパスワード (「アプリケーション シークレット」とも呼びます) が示されたポップアップ ダイアログが開きます。 *このパスワードをすぐにコピーして、アプリケーション ID と共に保存します。* これは、この後の手順で必要になります。 その後で、ダイアログを閉じます。

1. **[プラットフォーム]** セクションで、**[プラットフォームの追加]** をクリックします。

1. 開いたダイアログで、**[Web API]** を選択します。

1. **[アプリケーション ID URI]** が、"api://{App ID GUID}" という形式で生成されています。二重スラッシュと GUID の間に文字列 “localhost:44355/” を挿入します。全体の ID は、`api://localhost:44355/{App ID GUID}` のようになります。(**[アプリケーション ID URI]** の直後の **[スコープ]** 名のドメイン部分は、一致するように自動的に変更されます。これは、`api://localhost:44355/{App ID GUID}/access_as_user` のようになります)。

1. **[事前承認済みアプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。 次のそれぞれの ID を事前承認する必要があります。 1 つの ID を入力するたびに、新しい空のテキスト ボックスが表示されます。 (GUID のみを入力してください。)
 * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
 * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
 * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. それぞれの **[アプリケーション ID]** の横の **[スコープ]** ドロップダウンを開いて、`api://localhost:44355/{App ID GUID}/access_as_user` のボックスをオンにします。

1. **[プラットフォーム]** セクションの上部にある **[プラットフォームの追加]** を再度クリックして、**[Web]** を選択します。

1. **[プラットフォーム]** の下側の新しい **[Web]** セクションで、**[リダイレクト URL]** として `https://localhost:44355` を入力します。

    > 注:この記事の執筆時には、**[プラットフォーム]** セクションに **[Web API]** が表示されないことがあります。特に、**Web** プラットフォームを追加して*登録ページを保存*した後でページが最新の情報に更新されたときに発生します。**[Web API]** プラットフォームが登録に含まれていることを再確認する場合は、ページの下側にある **[アプリケーション マニフェストの編集]** ボタンをクリックします。マニフェストの **identifierUris** プロパティに、文字列 `api://localhost:44355/{App ID GUID}` が表示されている必要があります。また、値 `access_as_user` を保持している **value** サブプロパティがある **oauth2Permissions** プロパティもあります。

1. **[Microsoft Graph のアクセス許可]** セクションを下にスクロールして、**[委任されたアクセス許可]** サブセクションを表示します。**[追加]** ボタンを使用して、**[アクセス許可の選択]** ダイアログを開きます。

1. このダイアログ ボックスで、次に示すアクセス許可に対応するボックスをオンにします (既定でオンになっているものもあります)。 実際にアドイン自体に必要なのは最初のものだけですが、サーバー側コードで使用される MSAL ライブラリで `offline_access` および `openid` が必要とされます。 Office ホストがアドインの Web アプリケーションに対してトークンを取得するために、`profile` のアクセス許可が必要です。
 * Files.Read.All
 * offline_access
 * openid
 * profile

1. ダイアログの下部にある **[OK]** をクリックします。

1. 登録ページの下部にある **[保存]** をクリックします。

## <a name="grant-admin-consent-to-the-add-in"></a>アドインに管理者の同意を付与する

> **メモ:**この手順は、アドインの開発時にのみ必要になります。 実際に運用するアドインを Office ストアまたはアドイン カタログに展開した場合、インストール時に、各ユーザーが個別にそのアドインを信頼するか、管理者が組織のために同意することになります。

1. Visual Studio でアドインを実行していない場合は、**F5** キーを押して実行します。 この手順をスムーズに完了するには、IIS で実行する必要があります。

1. 次に示す文字列内のプレースホルダー “{application_ID}” は、アドインの登録時にコピーしたアプリケーション ID に置き換えます: `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. そうしてできた URL をブラウザーのアドレス バーに貼り付けて、そこに移動します。

1. ダイアログが表示されたら、管理者の資格情報を使用して Office 365 テナントにサインインします。

1. その後で、Microsoft Graph データにアクセスするためのアクセス許可をアドインに付与するように求めるダイアログが表示されます。**[承諾]** をクリックします。

1. ブラウザー ウィンドウ (タブ) は、アドインの登録時に指定した **[リダイレクト URL]** にリダイレクトされ、アドインのホーム ページがブラウザーで開かれます。

2. ブラウザーのアドレスバーには、GUID 値の付いた "tenant" クエリ パラメーターが表示されます。 これは、Office 365 テナントの ID です。 この値をコピーして保存します。 これは後の手順で使用します。

3. ウィンドウ (タブ) を閉じます。

1. Visual Studio のデバッガーを停止します。

## <a name="configure-the-add-in"></a>アドインを構成する

1. 次に示す文字列内のプレースホルダー "{tenant_ID}" は、前の手順で取得した Office 365 テナント ID に置き換えます。何らかの理由で、まだ ID を取得していない場合は、「[Office 365 テナント ID を検索する](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b)」に示したいずれかの方法を使用して ID を取得します。

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. Visual Studio で、web.config を開きます。**[appSettings]** セクションには、値を割り当てる必要のあるいくつかのキーがあります。

1. "ida:Issuer" という名前のキーの値として、手順 1 で作成した文字列を使用します。この値に、空白スペースが含まれていないことを確認してください。

1. 次に示す値を対応するキーに代入します。

|キー|値|
|:-----|:-----|
|ida:ClientID|アドインの登録時に取得したアプリケーション ID。|
|ida:Audience|アドインの登録時に取得したアプリケーション ID。|
|ida:Password|アドインの登録時に取得したパスワード。|


次に、4 つのキーの変更後の例を示します。 *ClientID と Audience が同一となっている点に注目してください*。 両方の目的に単一のキーを使用することもできますが、いつも同一となるとは限らないため、別々に保持しておくと、web.config のマークアップはより再利用しやすくなります。 また、別々のキーを使用することで、アドインが Office ホストに関連する OAuth リソースであり、かつ Microsoft Graph に関連する OAuth クライアントでもあるという概念が強調されます。

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    ```

> **注:**その他の **[appSettings]** セクションの設定は、未変更のままにします。

1. ファイルを保存して閉じます。

1. アドイン プロジェクトで、アドイン マニフェスト ファイル "Office-Add-in-ASPNET-SSO.xml" を開きます。

1. ファイルの最後までスクロールします。

1. `</VersionOverrides>` 終了タグの直前に、以下のマークアップがあります。

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}<Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. このマークアップ内の*両方の場所の*プレースホルダー “{application_GUID here}” を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。 「{」と「}」は ID の一部ではないので、これらを含めないでください。 これは、web.config の ClientID と Audience に使用したものと同じ ID です。

    > **メモ**:
    >* **[リソース]** の値は、アドインの登録に Web API プラットフォームを追加したときに設定した **[アプリケーション ID URI]** です。
    >* **[範囲]** セクションは、アドインが Office Store から販売された場合に、同意ダイアログ ボックスを生成するためにのみ使用します。

1. Visual Studio で、**[エラー一覧]** の **[警告]** タブを開きます。 `<WebApplicationInfo>` が `<VersionOverrides>` の有効な子ではないという警告が表示される場合は、Visual Studio 2017 プレビューのバージョンで SSO マークアップが認識されていません。 回避策として、Word、Excel、または PowerPoint のアドインに対して、次の操作を行います。 (Outlook アドインを使用している場合は、以下の回避策を参照してください。)

   - **Word、Excel、および PowerPoint の回避策**

   > 1. マニフェストの `</VersionOverrides>` の終了タグの直前の `<WebApplicationInfo>` セクションをコメント アウトします。

   > 2. F5 キーを押してデバッグ セッションを開始します。これにより、次のフォルダーにマニフェストのコピーが作成されます (これには、Visual Studio よりも**ファイル エクスプ ローラー**の方が容易にアクセスできます): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

   > 3. マニフェストのコピーから、`<WebApplicationInfo>` セクションの周囲のコメント構文を削除します。

   > 4. マニフェストのコピーを保存します。

   > 5. そして、次回 F5 キーを押したときに、Visual Studio がマニフェストのコピーを上書きしないようにする必要があります。 **ソリューション エクスプローラー**の最上部にあるソリューション ノード (どちらのプロジェクト ノードでもない) を右クリックします。

   > 6. コンテキスト メニューから **[プロパティ]** を選択します。**[ソリューション プロパティ ページ]** ダイアログ ボックスが開きます。

   > 7. **[構成プロパティ]** を展開し、**[構成]** を選択します。

   > 8. **Office-Add-in-ASPNET-SSO** プロジェクト (**Office-Add-in-ASPNET-SSO-WebAPI** プロジェクトでは*ありません*) の行で、**[ビルド]** と **[展開]** を選択解除します。

   > 9. **[OK]** をクリックしてダイアログ ボックスを閉じます。

   - **Outlook の回避策**

   > 1. 開発用コンピューターで、既存の `MailAppVersionOverridesV1_1.xsd` を探します。 `./Xml/Schemas/{lcid}` の下の Visual Studio インストール ディレクトリに配置されています。 たとえば、英語版 (米国) の VS 2017 32 ビットの標準インストールの場合、完全なパスは、`C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033` になります。

   > 2. 既存のファイルの名前を、`MailAppVersionOverridesV1_1.old` に変更します。

   > 3. 変更したこのファイルを、フォルダーにコピーします。[変更済みの MailAppVersionOverrides スキーマ](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. Visual Studio でメインのマニフェスト ファイルを保存して閉じます。

## <a name="code-the-client-side"></a>クライアント側のコードの作成

1. **[Scripts]** フォルダー内の Home.js ファイルを開きます。これには、一部のコードが既に含まれています。
    * `Office.initialize` メソッドへの割り当てが、`getGraphAccessTokenButton` ボタン クリック イベントへのハンドラーの割り当てになります。
    * `showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。

1. `Office.initialize` への割り当ての下に、次に示すコードを追加します。このコードについては、以下に注意してください。

    * `getAccessTokenAsync` は Office.js の新しい API です。これにより、アドインは Office ホスト アプリケーション (Excel、PowerPoint、Word など) に、アドインへのアクセス トークン (Office にサインインしているユーザーのトークン) を要求できるようになります。その Office ホスト アプリケーションが、Azure AD 2 エンドポイントにトークンを要求します。アドインの登録時に、アドインに対する Office ホストを事前認証しているため、Azure AD はトークンを送信します。
    * Office にサインインしているユーザーがいない場合、Office ホストはユーザーにサインインを求めるダイアログを表示します。
    * オプションのパラメーター `forceConsent` を false に設定すると、Office ホストにアドインへのアクセスを付与するための同意を求めるダイアログが表示されなくなります。

    ```js
    function getOneDriveFiles() {
        getDataWithToken({ forceConsent: false });
    }

    function getDataWithToken(options) {
        Office.context.auth.getAccessTokenAsync(options,
            function (result) {
                if (result.status === "succeeded") {
                    TODO1: Use the access token to get Microsoft Graph data.
                }
                else {
                    console.log("Code: " + result.error.code);
                    console.log("Message: " + result.error.message);
                    console.log("name: " + result.error.name);
                    document.getElementById("getGraphAccessTokenButton").disabled = true;
                }
            });
    }
    ```

1. TODO1 を次に示す行に置き換えます。`getData` メソッドとサーバー側の "/api/values" ルートは、この後の手順で作成します。エンドポイントには、相対 URL を使用します。これは、その URL がアドインと同じドメインでホストされている必要があるためです。

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. `getOneDriveFiles` メソッドの下に、次を追加します。このユーティリティ メソッドは、特定の Web API エンドポイントを呼び出して、Office ホスト アプリケーションがアドインへのアクセスに使用したものと同じアクセス トークンを渡します。サーバー側では、このアクセス トークンが Microsoft Graph へのアクセス トークンを取得するための「代理」フローで使用されます。

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            TODO2: Handle errors and the case where Microsoft Graph
                   requires additional form of authentication.
        });
    }
    ```

1. TODO2 を以下に示すコードに置き換えます。 このコードについては、次の点に注意してください。

    * 失敗の原因が、Microsoft Graph が認証の追加のフォームを要求したためだった場合、`exceptionMessage` は "capolids" を含む JSON 文字列になります。 その場合、Office ホストは、新しいトークンを取得する必要があります。  
    * Office ホストに渡されるべき例外メッセージは、すべての必要な認証のフォームについてユーザー入力を求めるように AAD に指示します。Office ホストは新しいトークンを要求するとき、AAD にそれを渡します。
    * `authChallenge` オプションは、Office ホストにこの文字列を渡すメソッドです。
    * エラーが追加の認証の要求以外の場合は、コンソールに記録されます。

    ```js
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    if (exceptionMessage.indexOf("capolids") !== -1) {
        getDataWithToken({ authChallenge: exceptionMessage });
    } else {
        console.log(result.error);
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

    ```
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

    ```
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. TODO3 を次のように置き換えます。 メモ:

    * このコードでは、Office ホストから得られるアクセス トークン (`getData` のクライアント側呼び出しによって渡されるトークン) で指定された対象ユーザーとトークン発行者が web.config で指定された値と一致する必要があることを OWIN に指示します。
    * `SaveSigninToken` を `true` に設定することで、OWIN は Office ホストからの Raw トークンを保存するようになります。これは、アドインが「代理」フローで Microsoft Graph へのアクセス トークンを取得するために必要になります。
    * OWIN ミドルウェアでは、スコープは検証されません。`access_as_user` が含まれている必要があるアクセス トークンのスコープは、コントローラーで検証されます。

    ```
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. TODO4 を次のように置き換えます。 メモ:

    * より一般的な `UseWindowsAzureActiveDirectoryBearerAuthentication` は Azure AD V2 エンドポイントに準拠していないため、その代わりとしてメソッド `UseOAuthBearerAuthentication` が呼び出されます。
    * このメソッドに渡される探索 URL は、Office ホストから受け取ったアクセス トークンの署名の検証に必要になるキーを取得するための方法を OWIN ミドルウェアが取得する場所になります。

    ```
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
            {
                AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
            });
    ```

1. ファイルを保存して閉じます。

### <a name="create-the-apivalues-controller"></a>/api/values コントローラーを作成する

1. ファイル **Controllers\ValueController.cs** を開きます。

2. ファイルの先頭に、次に示す `using` ステートメントがあることを確認します。

    ```
    using Microsoft.Identity.Client;
    using System;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web;
    using System.Web.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. `ValuesController` を宣言している行のすぐ上に、`[Authorize]` 属性を追加します。 これにより、コントローラーのメソッドが呼び出されるたびに、アドインが先程の手順で構成した認証プロセスを実行するようにします。 アドインに対し有効なアクセス トークンを持つ呼び出し元だけが、コントローラーのメソッドを呼び出すことができます。

4. 次に示すメソッドを `ValuesController` に追加します。

    ```
    // GET api/values
    public async Task<IEnumerable<string>> Get()
    {
        // TODO5: Validate the scopes of the access token.
    }
    ```

5. TODO5 を次に示すコードに置き換えます。このコードでは、`access_as_user` を含むトークンで指定されているスコープを検証します。

    ```
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO6: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO7: Get the access token for Microsoft Graph.
        // TODO8: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO9: Remove excess information from the data and send the data to the client.
    }
    return new string[] { "Error", "Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user." };
    ```

> **注:**`access_as_user` スコープだけを使用して、Office アドインの代理フローを処理する API を承認する必要があります。サービス内の他の API は、独自のスコープ要件が必要です。 これにより、Office が取得するトークンでアクセスできるものが制限されます。

6. TODO6 を次のコードで置き換えます。メモ:
    * このコードでは、Office ホストから受け取った Raw アクセス トークンを別のメソッドに渡される `UserAssertion` オブジェクトに変換します。
    * アドインは、Office ホストとユーザーがアクセスする必要のあるリソース (または対象ユーザー) の役割を果たさなくなります。この時点で、それ自体が Microsoft Graph にアクセスする必要があるクライアントになります。`ConfidentialClientApplication` は MSAL の「クライアント コンテキスト」オブジェクトになります。
    * `ConfidentialClientApplication` コンストラクターへの 3 番目のパラメーターはリダイレクト URL です。これは、実際には「代理」フローで使用されることはありませんが、正しい URL を使用することをお勧めします。4 番目と 5 番目のパラメーターは、永続ストアを定義するために使用できます。このストアにより、有効期限が切れていないトークンをアドインの異なるセッション間で再使用できるようになります。このサンプルでは、永続ストアは実装していません。
    * MSAL では `openid`、`offline_access` の各スコープが機能することが必要ですが、コードがこれらを重複して要求するとエラーがスローされます。 コードが `profile` を要求した場合にもエラーがスローされます。それは、実際には Office ホスト アプリケーションがアドインの Web アプリケーションに対しトークンを取得するときだけに使用します。 そのため、`Files.Read.All` のみが明示的に要求されます。

    ```
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. TODO7 を以下のコードに置き換えます。 メモ:

    * `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` メソッドは、最初にメモリ内の MSAL キャッシュで一致するアクセス トークンを探します。それが見つからなかった場合にのみ、Azure AD V2 エンドポイントとの「代理」フローを開始します。
    * MS Graph リソースが多要素認証を必要とし、ユーザーがまだそれを提供していない場合、AAD は Claims プロパティが含まれている例外をスローします。
    * Claims プロパティの値は、クライアントに渡されなければなりません。そしてクライアントは、それを Office ホストに渡します。Office ホストは、新しいトークンの要求にそれを入れます。 AAD は、認証のすべての要求されたフォームをユーザーに示します。
    * `MsalUiRequiredException` タイプではない例外はすべて、意図的にキャッチされないため、クライアントまで伝達されます。

    ```
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalUiRequiredException e)
    {        
        if (String.IsNullOrEmpty(e.Claims))
        {
            throw e;
        }
        else
        {
            throw new HttpException(e.Claims);
        }   
    }
    ```

8. TODO8 を次のように置き換えます。 メモ:

    * `GraphApiHelper` クラスと `ODataHelper` クラスは、**[Helpers]** フォルダー内のファイルで定義されています。`OneDriveItem` クラスは、**[Models]** フォルダー内のファイルで定義されています。これらのクラスについての詳しい説明は、承認や SSO に関連していないため、この記事の対象外になります。
    * 実際に必要なデータのみを Microsoft Graph に要求することでパフォーマンスが向上します。そのため、このコードでは、` $select` クエリ パラメーターで name プロパティのみが必要なことを指定し、`$top` パラメーターで最初の 3 つのフォルダー名またはファイル名のみが必要なことを指定しています。

    ```
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    var getFilesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    ```

9. TODO9 を次のように置き換えます。 上記のコードでは OneDrive アイテムの *name* プロパティのみを要求していますが、Microsoft Graph は常に OneDrive アイテムの *eTag* プロパティを含めます。 クライアントに送信するペイロードを縮小するために、次に示すコードでは結果をアイテム名のみで再構築しています。

    ```
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in getFilesResult)
    {
      itemNames.Add(item.Name);
    }                    
    return itemNames;
    ```

## <a name="run-the-add-in"></a>アドインを実行する

1. 結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。

1. Visual Studio で、F5 キーを押します。PowerPoint が開き、**[ホーム]** リボンに **[SSO ASP.NET]** グループが表示されます。

1. このグループ内の **[アドインの表示]** ボタンをクリックすると、作業ウィンドウにアドインの UI が表示されます。

1. **[OneDrive からファイルを取得]** ボタンをクリックします。 Office にサインインしていない場合は、サインインを求めるダイアログが表示されます。
    > **メモ:**以前に別の ID で Office にサインオンしていて、そのときに開いたいくつかの Office アプリケーションが引き続き開いている場合、Office がその ID を確実に変更するとは限りません (PowerPoint で ID が変更済みのように表示されている場合でも)。 このような場合は、Microsoft Graph への呼び出しが失敗するか、以前の ID からのデータが返される可能性があります。 これを防止するには、必ず*他のすべての Office アプリケーションを閉じて*から、**[OneDrive からファイルを取得]** を押します。

1. サインインすると、ボタンの下に OneDrive のファイルとフォルダーのリストが表示されます。 これには、15 秒以上かかることがあります (特に初回実行時)。

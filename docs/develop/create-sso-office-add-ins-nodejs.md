# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>シングル サインオンを使用する Node.js Office アドインを作成する

ユーザーは、このサインイン プロセスを利用してユーザーを承認する Office および Office Web アドインにサインインできます。こうして承認されたユーザーは、アドインと Microsoft Graph への 2 度目のサインオンの必要がなくなります。概要については、「[Office アドインで SSO を有効化する](../../docs/develop/sso-in-office-add-ins.md)」を参照してください。

この記事では、Node.js と express を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。 

> **注:**ASP.NET ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する ASP.NET Office アドインの作成](../../docs/develop/create-sso-office-add-ins-aspnet.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

* [Node および npm](https://nodejs.org/en/)、バージョン 6.9.4 以降。
* [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント。)
* TypeScript バージョン 2.2.2 以降。
* Office 2016 バージョン 1708、ビルド 8424.nnnn 以降 (「クイック実行」と呼ばれることもある Office 365 のサブスクリプション バージョン)。このバージョンを入手するには、Office Insider への参加が必要になることがあります。詳細については、「[Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1)」を参照してください。

## <a name="set-up-the-starter-project"></a>スタート プロジェクトをセットアップする

1. 「[Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso)」にあるリポジトリを複製するかダウンロードします。 


    > **注:**サンプルには 2 つのバージョンがあります。 
    > 
    > * **[Before]** フォルダーはスタート プロジェクトです。SSO や承認に直接関連しない UI などの側面は、既に完了しています。この記事で後述する各セクションでは、これを完成させるための手順を順に説明します。 
    > * このサンプルの **[Completed]** バージョンは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。完成済みのバージョンを使用する場合は、この記事の手順をそのまま実行しますが、[Before] を [Completed] に置き換えて、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションを省略してください。

1. **[Before]** フォルダー内で Git bash コンソールを開きます。

2. コンソールで `npm install` を入力して、package.json ファイル内のアイテム化されたすべての依存関係をインストールします。

3. コンソールで `npm run build ` を入力して、プロジェクトをビルドします。 
     > 注:いくつかの使用されていない変数が宣言されているという、ビルド エラーが発生することがあります。 これらのエラーは無視してください。 これらは、後で追加する一部のコードが見つからないという「Before」バージョンのサンプルの副作用です。

## <a name="register-the-add-in-with-azure-ad-v2-endpoint"></a>Azure AD V2 エンドポイントにアドインを登録する

1. [https://apps.dev.microsoft.com](https://apps.dev.microsoft.com) に移動します。 

1. 管理者の資格情報を使用して Office 365 テナントにサインインします。たとえば、MyName@contoso.onmicrosoft.com

1. **[アプリの追加]** をクリックします。

1. ダイアログが表示されたら、アプリ名として「Office-Add-in-NodeJS-SSO」を使用して、**[アプリケーションの作成]** をクリックします。

1. アプリの構成ページが開いたら、**[アプリケーション ID]** をコピーして保存します。これは、この後の手順で使用します。 

    > 注:この ID は、Office ホスト アプリケーション (たとえば、PowerPoint、Word、Excel) などの別のアプリケーションが、このアプリケーションへの承認されたアクセスを求めるときの「対象ユーザー」値になります。また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。

1. **[アプリケーション シークレット]** セクションで、**[新しいパスワードを生成]** をクリックします。新しいパスワード (「アプリケーション シークレット」とも呼びます) が示されたポップアップ ダイアログが開きます。*このパスワードをすぐにコピーして、アプリケーション ID と共に保存します。*これは、この後の手順で必要になります。その後で、ダイアログを閉じます。

1. **[プラットフォーム]** セクションで、**[プラットフォームの追加]** をクリックします。 

1. 開いたダイアログで、**[Web API]** を選択します。

1. **[アプリケーション ID URI]** が、"api://{App ID GUID}" という形式で生成されています。二重スラッシュと GUID の間に文字列 “localhost:3000” を挿入します。全体の ID は、`api://localhost:3000/{App ID GUID}` のようになります。(**[アプリケーション ID URI]** の直後の **[スコープ]** 名のドメイン部分は、一致するように自動的に変更されます。これは、`api://localhost:3000/{App ID GUID}/access_as_user` のようになります)。

1. この手順と次の手順で、Office アプリケーションにアドインへのアクセス権を付与します。 **[事前承認済みアプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。 次のそれぞれの ID を事前承認する必要があります。 1 つの ID を入力するたびに、新しい空のテキスト ボックスが表示されます。 (GUID のみを入力してください。)

 * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
 * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
 * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online) 

1. それぞれの **[アプリケーション ID]** の横の **[スコープ]** ドロップ ダウンを開いて、`api://localhost:44355/{App ID GUID}/access_as_user` のボックスをオンにします。

1. **[プラットフォーム]** セクションの上部にある **[プラットフォームの追加]** を再度クリックして、**[Web]** を選択します。

1. **[プラットフォーム]** の下側の新しい **[Web]** セクションで、**[リダイレクト URL]** として `https://localhost:3000` を入力します。 

    > 注:この記事の執筆時には、**[プラットフォーム]** セクションに **[Web API]** が表示されないことがあります。特に、**Web** プラットフォームを追加して*登録ページを保存*した後でページが最新の情報に更新されたときに発生します。**[Web API]** プラットフォームが登録に含まれていることを再確認する場合は、ページの下側にある **[アプリケーション マニフェストの編集]** ボタンをクリックします。マニフェストの **identifierUris** プロパティに、文字列 `api://localhost:3000/{App ID GUID}` が表示されている必要があります。また、値 `access_as_user` を保持している **value** サブプロパティがある **oauth2Permissions** プロパティもあります。

1. **[Microsoft Graph のアクセス許可]** セクションを下にスクロールして、**[委任されたアクセス許可]** サブセクションを表示します。**[追加]** ボタンを使用して、**[アクセス許可の選択]** ダイアログを開きます。

1. ダイアログ ボックスで、次の各アクセス許可のボックスをオンにします。 
    * Files.Read.All
    * profile

1. ダイアログの下側にある **[OK]** をクリックします。

1. 登録ページの下側にある **[保存]** をクリックします。

## <a name="grant-admin-consent-to-the-add-in"></a>アドインに管理者の同意を付与する

> **メモ:**この手順は、アドインの開発時にのみ必要になります。 実際に運用するアドインを Office ストアまたはアドイン カタログに展開する場合、各ユーザーはそのアドインをインストールする際にそのアドインを個別に信頼します。

1. 次に示す文字列内のプレースホルダー "{application_ID}" は、アドインの登録時にコピーしたアプリケーション ID に置き換えます。

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. その結果の URL をブラウザーのアドレス バーに貼り付けて、その場所に移動します。

1. ダイアログが表示されたら、管理者の資格情報を使用して Office 365 テナントにサインインします。

1. その後で、Microsoft Graph データにアクセスするためのアクセス許可をアドインに付与するように求めるダイアログが表示されます。**[承諾]** をクリックします。 

1. ブラウザー ウィンドウ/タブが、アドインの登録時に指定した**リダイレクト URL** にリダイレクトされ、アドインが実行中の場合は、アドインのホーム ページがブラウザーで開かれます。 アドインが実行中でない場合は、localhost:3000 でリソースが見つからないか開けないというエラーが表示されます。 *ただし、リダイレクションが試行されたということが、管理者の同意プロセスが正常に完了したことを意味しています*。 そのため、ホーム ページが開かれたか、このエラーが発生したかに関係なく、次の手順に進むことができます。

2. ブラウザーのアドレスバーには、GUID 値の付いた "tenant" クエリ パラメーターが表示されます。これは、Office 365 テナントの ID です。この値をコピーして保存します。これは後の手順で使用します。

3. ウィンドウ (タブ) を閉じます。

## <a name="configure-the-add-in"></a>アドインを構成する

1. コード エディターで、src\server.ts ファイルを開きます。先頭近くに、`AuthModule` クラスのコンストラクターの呼び出しがあります。コンストラクターには、値を割り当てる必要がある、文字列のパラメーターがあります。

2. `client_id` プロパティの場合は、アドインの登録時に保存したアプリケーション ID でプレースホルダーの `{client GUID}` を置き換えます。完了すると、単一引用符で囲まれた GUID のみになります。"{}" 文字は取り去る必要があります。

3. `client_secret` プロパティの場合は、アドインの登録時に保存したアプリケーション シークレットでプレースホルダーの `{client secret}` を置き換えます。

4. `audience` プロパティの場合は、アドインの登録時に保存したアプリケーション ID でプレースホルダーの `{audience GUID}` を置き換えます。(`client_id` プロパティに割り当てた値とまったく同じになります)。
  
3. `issuer` プロパティに割り当てた文字列には、プレースホルダーの *{O365 tenant GUID}* があります。これは、前の最後の手順で保存した Office 365 テナント ID で置き換えます。何らかの理由で、まだ ID を取得していない場合は、「[Office 365 テナント ID を検索する](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b)」に示したいずれかの方法を使用して ID を取得します。完了すると、`issuer` プロパティの値は、次に示すようなものになります。

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. `AuthModule` コンストラクターのその他の値は未変更のままにしておきます。 ファイルを保存して閉じます。

1. プロジェクトのルートにある、アドイン マニフェスト ファイル「Office-Add-in-NodeJS-SSO.xml」を開きます。

1. ファイルの最後までスクロールします。

1. 最後の `</VersionOverrides>` タグの直前に、次に示すマークアップが見つかります。

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. このマークアップ内の*両方の場所の*プレースホルダー “{application_GUID here}” を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。 (「{」と「}」は ID の一部ではないので、これらを含めないでください。)。これは、web.config の ClientID と Audience に使用したものと同じ ID です。

    >注意: 
    >
    >* **[リソース]** の値は、アドインの登録に Web API プラットフォームを追加したときに設定した **[アプリケーション ID URI]** です。
    >* **[範囲]** セクションは、アドインが Office Store から販売された場合に、同意ダイアログ ボックスを生成するためにのみ使用します。

1. ファイルを保存して閉じます。

## <a name="code-the-client-side"></a>クライアント側のコードを作成する

1. **[public]** フォルダー内の program.js ファイルを開きます。これには、一部のコードが既に含まれています。

    * `Office.initialize` メソッドへの割り当てが、`getGraphAccessTokenButton` ボタン クリック イベントへのハンドラーの割り当てになります。
    * `showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。

1. `Office.initialize` への割り当ての下に、次に示すコードを追加します。このコードについては、以下に注意してください。 

    * 代理 (on-behalf-of) フローを使用するための `getDataWithoutAuthChallenge` 関数が最初の試行で呼び出されます。 必要なのは単一要素認証だけであることを前提とします。 多要素認証が必要な場合には、後の手順でコードを追加して対応します。
    * `getAccessTokenAsync` は Office.js の新しい API です。これにより、アドインは Office ホスト アプリケーション (Excel、PowerPoint、Word など) に、アドインへのアクセス トークン (Office にサインインしているユーザーのトークン) を要求できるようになります。 その Office ホスト アプリケーションが、Azure AD 2.0 エンドポイントにこのトークンを要求します。 アドインの登録時に、アドインに対する Office ホストを事前認証しているため、Azure AD はそのトークンを送信します。 
     * Office にサインインしているユーザーがいない場合、Office ホストはユーザーにサインインを求めるダイアログを表示します。 
     * オプションのパラメーター `forceConsent` を false に設定すると、Office ホストにアドインへのアクセスを付与するための同意を求めるダイアログが表示されなくなります。

    ```js
    function getOneDriveItems() {
        getDataWithoutAuthChallenge();
    }   
    
    function getDataWithoutAuthChallenge() {       
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    // TODO1: Use the access token to get Microsoft Graph data.
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

1. TODO1 を次に示す行に置き換えます。 `getData` メソッドとサーバー側の “/api/onedriveitems” ルートを、この後の手順で作成します。 エンドポイントには、相対 URL を使用します。これは、その URL がアドインと同じドメインでホストされている必要があるためです。

    ```
    accessToken = result.value;
    getData("/api/onedriveitems", accessToken);
    ```

1. `getOneDriveFiles` メソッドの下に、次を追加します。このユーティリティ メソッドは、特定の Web API エンドポイントを呼び出して、Office ホスト アプリケーションがアドインへのアクセスに使用したものと同じアクセス トークンを渡します。サーバー側では、このアクセス トークンが Microsoft Graph へのアクセス トークンを取得するための「代理」フローで使用されます。 

    ```
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            TODO2: Display data and handle demand for multi-factor authentication.
        })
        .fail(function (result) {
            console.log(result.error);
       });
    }
    ```

1. TODO2 を次に示すコードに置き換えます。このコードに関する注意:
    * Microsoft Graph のターゲットが追加の認証要素を要求する場合、結果はデータになりません。 その結果は、ユーザーが提供する必要がある追加の要素を AAD に通知する Claims JSON になります。 その場合にクライアントは、AAD が必要なダイアログを表示するように、この Claims の文字列を AAD に渡す新しいサインオンを開始する必要があります。
    * 結果が Claims JSON の場合、それには文字列の "capolids" が含まれます。
    * 後の手順で `getDataUsingAuthChallenge` 関数を作成します。

    ```
    if (result[0].indexOf('capolids') !== -1) {                
        result[0] = JSON.parse(result[0])
        getDataUsingAuthChallenge(result[0]);
    } else {  
        showResult(result);
    }
    ```

1. ファイルの `getData` 関数のすぐ下に、次の関数を追加します。 この関数については、次の点に注意してください。
    * この関数は、AAD が追加の認証要素を要求した場合に使用されます。 
    * この関数は、追加の認証要素を提供するようにユーザーに求めるダイアログが表示される 2 番目のサインオンをトリガーします。 
    * `authChallenge` オプションには、AAD が入力を求めるダイアログを表示すべき要素を AAD に通知する文字列が含まれています。 Office ホストは、アドインにアドイン トークンを要求する際、この文字列を AAD に渡します。

    ```
    function getDataUsingAuthChallenge(authChallengeString) {       
        Office.context.auth.getAccessTokenAsync({authChallenge: authChallengeString},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData("/api/onedriveitems", accessToken);
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

1. ファイルを保存して閉じます。

## <a name="code-the-server-side"></a>サーバー側のコードを作成する

変更の必要があるサーバー側のファイルは 2 つあります。 
- src\auth.js では、承認のヘルパー関数を提供します。これには、各種の承認フローで使用される汎用のメンバーが既に含まれています。これには、「代理」フローを実装するための関数を追加する必要があります。
- src\server.js ファイルには、サーバーと express ミドルウェアを実行するために必要な基本的なメンバーが含まれています。これには、ホーム ページと Microsoft Graph データを取得するための Web API を提供する関数を追加する必要があります。

### <a name="create-a-method-to-exchange-tokens"></a>トークンを交換するためのメソッドを作成する

1. \src\auth.ts ファイルを開きます。`AuthModule` クラスに、次に示すメソッドを追加します。このコードについては、以下に注意してください。
    * jwt パラメーターは、アプリケーションへのアクセス トークンです。「代理」フローでは、これはリソースへのアクセス トークンの AAD と交換されます。
    * scopes パラメーターには既定の値がありますが、このサンプルではコード呼び出しによってオーバーライドしています。
    * resource パラメーターは省略可能です。STS が AAD V2 エンドポイントの場合は使用しないでください。後者の場合は scopes から resource 推測され、resource が HTTP 要求で送信されるとエラーを返します。 
    

    ```
    private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        try {
            // TODO3: Construct the parameters that will be sent in the body of the 
            //        HTTP Request to the STS that starts the "on behalf of" flow.
            // TODO4: Send the request to the STS.
            // TODO5: Process the response and persist the access token to resource.
        }
        catch (exception) {
            throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                        + JSON.stringify(exception), 
                                        exception);
        }
    }
    ```

2. TODO3 を次に示すコードに置き換えます。 このコードについては、次の点に注意してください。
    * 「代理」ワークフローをサポートする STS は、HTTP 要求の本文に特定のプロパティ/値ペアが含まれていることを期待します。このコードは、要求の本文になるオブジェクトを構築します。 
    * resource プロパティは、リソースがメソッドに渡された場合にのみ本文に追加されます。

    ```
    const v2Params = {
            client_id: this.clientId,
            client_secret: this.clientSecret,
            grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
            assertion: jwt,
            requested_token_use: 'on_behalf_of',
            scope: scopes.join(' ')
        };
        let finalParams = {};
        if (resource) {
            // In JavaScript we could just add the resource property to the v2Params
            // object, but that won't compile in TypeScript.
            let v1Params  = { resource: resource };  
            for(var key in v2Params) { v1Params[key] = v2Params[key]; }
            finalParams = v1Params;
        } else {
            finalParams = v2Params;
        } 
    ```

3. TODO4 を次に示すコードに置き換えます。このコードでは、HTTP 要求を STS のトークン エンドポイントに送信します。

    ```
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. TODO5 を次に示すコードに置き換えます。 このコードはリソースへのアクセス トークンを永続化して、有効期限になると、そのアクセス トークンを返します。 コードを呼び出すことで、期限切れになっていないリソースへのアクセス トークンが再使用されるため、STS への不要な呼び出しを回避できます。 この動作のしくみは、次のセクションで説明します。

    ```
    if (res.status !== 200) {
        TODO6: Handle failure and the case where AAD asks for additional
               authentication factors.
    }
    const json = await res.json();
    // Persist the token and it's expiration time.
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

5. TODO6 を次に示すコードに置き換えます。 このコードについては、次の点に注意してください。

    * ユーザーがパスワードだけで Office にサインオンできる場合でも、Microsoft Graph のいくつかのターゲット (たとえば、OneDrive) にアクセスするために、追加の認証要素を提供するようにユーザーに要求する、Azure Active Directory の構成があります。 この場合、AAD は `Claims` プロパティが含まれた応答を送信します。 
    * この `Claims` の値をクライアントに戻して、そのユーザーの 2 番目のサインオンを開始し、AAD への呼び出しに `Claims` の値を含めるようにする必要があります。 AAD は、追加の認証要素を提供するようにユーザーに求めるダイアログを表示します。
    * コードは、予防措置としてユーザーがパスワードだけでログインしたときに取得されたアクセス トークンすべてのキャッシュをクリアします。  

    ```
    const exception = await res.json();
    // Check if AAD is the STS.
    if (this.stsDomain === 'https://login.microsoftonline.com') {
        if (JSON.stringify(exception.claims)) {                       
            ServerStorage.clear();
            return JSON.stringify(exception.claims);    
        } else {                    
            throw exception;
        }
    }
    else {                    
        throw exception;
    }
    ```

5. ファイルを閉じないで保存します。

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>「代理」ワークフローを使用してリソースにアクセスするメソッドを作成する

1. 引き続き src/auth.ts で、次に示すメソッドを `AuthModule` クラスに追加します。このコードについては、以下に注意してください。
    * `exchangeForToken` メソッドへのパラメーターに関する上記のコメントは、このメソッドのパラメーターにも当てはまります。
    * このメソッドは最初に永続ストレージをチェックして、期限が切れておらず、次の 1 分間に期限が切れない、リソースへのアクセス トークンがないか調べます。 このメソッドは、直前のセクションで作成した `exchangeForToken` メソッドを呼び出す必要がある場合にのみ、そのメソッドを呼び出します。

    ```
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. ファイルを保存して閉じます。

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>アドインのホーム ページとデータを提供するエンドポイントを作成する

1. src\server.ts ファイルを開きます。 

2. 次に示すメソッドをファイルの末尾に追加します。このメソッドにより、アドインのホーム ページを提供します。アドイン マニフェストで、ホーム ページの URL を指定します。

    ```
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. ファイルの末尾に次のメソッドを追加します。このメソッドにより、`onedriveitems` API に対する要求を処理します。
    ```
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Send to the client only the data that it actually needs.
    })); 
    ```

4. TODO7 を次に示すコードに置き換えます。このコードでは、Office ホスト アプリケーションから受け取ったアクセス トークンを検証します。 `verifyJWT` メソッドは、src\auth.ts ファイルで定義されています。 このメソッドは、常に対象ユーザーと発行者を検証します。 省略可能なパラメーターを使用して、アクセス トークンのスコープが `access_as_user` であることを検証する必要もあることを指定します。 これは、「代理 (on behalf flow)」フローによって Microsoft Graph へのアクセストークンを取得するために、ユーザーと Office ホストが必要とする、アドインに対する唯一のアクセス許可です。 

    ```
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

> **注:**`access_as_user` スコープだけを使用して、Office アドインの代理フローを処理する API を承認する必要があります。サービス内の他の API は、独自のスコープ要件が必要です。 これにより、Office が取得するトークンでアクセスできるものが制限されます。

5. TODO8 を次のコードで置き換えます。 このコードについては、次の点に注意してください。

    * `acquireTokenOnBehalfOf` の呼び出しには、resource パラメーターは含まれません。これは、resource プロパティをサポートしていない AAD V2.0 エンドポイントで `AuthModule` オブジェクト (`auth`) を作成したためです。
    * この呼び出しの 2 番目のパラメーターでは、OneDrive 上のユーザーのファイルとフォルダーのリストを取得するために、アドインが必要とするアクセス許可を指定します。 (`profile` アクセス許可は要求されません。これは、このアクセス許可が、Microsoft Graph へのアクセス トークン用のトークンでやり取りしているときではなく、Office ホストがアドインへのアクセス トークンを取得するときにだけ必要であるためです。)
    * 応答が 'capolids" を含む文字列である場合、これは多要素認証が要求される、AAD からの claims (要求) メッセージです。 このメッセージはクライアントに渡され、クライアントはそのメッセージを使用して 2 番目のサインオンを開始します。 この文字列は、AAD がユーザーに提供を求めるダイアログを表示する追加の認証要素を AAD に指示します。

    ```
    let graphToken = null;
    const tokenAcquisitionResponse = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    if (tokenAcquisitionResponse.includes('capolids')) {
        const claims: string[] = [];
        claims.push(tokenAcquisitionResponse);
        return res.json(claims);
    } else {
        // The response is the token to Microsoft Graph itself. Rename it so remaining code
        // is self-documenting.
        graphToken = tokenAcquisitionResponse;
    }
    ```

6. TODO9 を次に示す行に置き換えます。 このコードについては、次の点に注意してください。

    * MSGraphHelper クラスは、src\msgraph-helper.ts で定義されています。 
    * 返す必要があるデータが最小になるように、name プロパティと最初の 3 つのアイテムのみが必要なことを指定しています。

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. TODO10 を次に示すコードに置き換えます。 Microsoft Graph は、`name` プロパティのみを要求した場合でも、アイテムごとに、いくつかの OData メタデータと 1 つの **eTag** プロパティを返す点に注意してください。 このコードでは、アイテムの名前のみをクライアントに送信します。

    ```
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. ファイルを保存して閉じます。

## <a name="deploy-the-add-in"></a>アドインを展開する

次に、Office がアドインを検索する場所を認識できるようにする必要があります。

1. ネットワーク共有を作成するか、[フォルダーをネットワークに共有します](https://technet.microsoft.com/en-us/library/cc770880.aspx)。

2. プロジェクトのルートから、Office-Add-in-NodeJS-SSO.xml マニフェスト ファイルのコピーを共有フォルダーに配置します。

3. PowerPoint を起動して、ドキュメントを開きます。

4. **[ファイル]** タブを選択して、**[オプション]** を選択します。

5. [**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。

6. **[信頼されているアドイン カタログ]** を選択します。

7. **[カタログの URL]** フィールドに、Office-Add-in-NodeJS-SSO.xml があるフォルダー共有へのネットワーク パスを入力して、**[カタログの追加]** を選択します。

8. **[メニューに表示する]** チェック ボックスをオンにして、**[OK]** を選択します。

9. これらの設定は Microsoft Office を次回起動したときに適用されることを示すメッセージが表示されます。PowerPoint を終了します。

## <a name="build-and-run-the-project"></a>プロジェクトのビルドと実行

プロジェクトのビルドと実行には 2 つの方法があり、Visual Studio Code を使用しているかどうかによって決まります。どちらの方法でも、プロジェクトをビルドして、コードに変更を加えたときには自動的に再ビルドしてから再実行します。

1. Visual Studio Code を使用していない場合: 
 1. ノード ターミナルを開いて、プロジェクトのルート フォルダーに移動します。
 2. ターミナルで、「**npm run build**」と入力します。 
 3. 2 番目のノード ターミナルを開いて、プロジェクトのルート フォルダーに移動します。
 4. ターミナルで、「**npm run start**」と入力します。

2. VS Code を使用している場合:
 1. VS Code でプロジェクトを開きます。
 2. CTRL + SHIFT + B を押して、プロジェクトをビルドします。
 3. F5 を押して、デバッグ セッションでプロジェクトを実行します。


## <a name="add-the-add-in-to-an-office-document"></a>Office ドキュメントにアドインを追加する

1. PowerPoint を再起動して、プレゼンテーションを開くか作成します。 

2. PowerPoint の **[開発]** タブで、**[個人用アドイン]** を選択します。

3. **[共有フォルダー]** タブを選択します。

4. **[SSO NodeJS Sample]** を選択して、**[OK]** を選択します。

5. **[ホーム]** リボンに、**[SSO NodeJS]** という新しいグループが表示され、**[アドインの表示]** というラベルの付いたボタンとアイコンが含まれています。 

## <a name="test-the-add-in"></a>アドインをテストする

1. 結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。

2. **[アドインの表示]** ボタンをクリックして、アドインを開きます。

2. [ようこそ] ページでアドインが開きます。 **[OneDrive からファイルを取得]** ボタンをクリックします。

2. Office にサインインしている場合は、このボタンの下に OneDrive にあるファイルとフォルダーのリストが表示されます。これは、初回実行時には 15 秒以上かかることがあります。

3. Office にサインインしていない場合は、ポップアップが表示され、サインインするように求められます。サインインが完了すると、数秒後にファイルとフォルダーが表示されます。*2 回目にボタンをクリックする必要はありません。*
> **メモ:**以前に別の ID で Office にサインオンしていて、そのときに開いたいくつかの Office アプリケーションが引き続き開いている場合、Office がその ID を確実に変更するとは限りません (PowerPoint で ID が変更済みのように表示されている場合でも)。 このような場合は、Microsoft Graph への呼び出しが失敗するか、以前の ID からのデータが返される可能性があります。 これを防止するには、必ず*他のすべての Office アプリケーションを閉じて*から、**[OneDrive からファイルを取得]** を押します。

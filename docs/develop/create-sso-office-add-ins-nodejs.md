---
title: シングル サインオンを使用する Node.js Office アドインを作成する
description: Office シングル サインオンを使用するNode.js ベースのアドインを作成する方法について説明します。
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 35128da43b3f27a58df5e188a5001bfa8aba4a4c
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/28/2022
ms.locfileid: "68841729"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>シングル サインオンを使用する Node.js Office アドインを作成する

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

この記事では、アドインでシングル サインオン (SSO) を有効にするプロセスについて説明します。 作成するサンプル アドインには 2 つの部分があります。Microsoft Excel に読み込まれる作業ウィンドウと、作業ウィンドウの Microsoft Graph への呼び出しを処理する中間層サーバー。 中間層サーバーは、Node.jsと Express を使用して構築され、 `/getuserfilenames`ユーザーの OneDrive フォルダー内の最初の 10 個のファイル名の一覧を返す 1 つの REST API を公開します。 作業ウィンドウでは、 メソッドを `getAccessToken()` 使用して、中間層サーバーにサインインしているユーザーのアクセス トークンを取得します。 中間層サーバーは、On-Behalf-Of フロー (OBO) を使用して、Microsoft Graph にアクセスできる新しいサーバーのアクセス トークンを交換します。 このパターンを拡張して、任意の Microsoft Graph データにアクセスできます。 作業ウィンドウは、Microsoft Graph サービスが必要な場合に、常に中間層 REST API を呼び出します (アクセス トークンを渡します)。 中間層は、OBO によって取得されたトークンを使用して Microsoft Graph サービスを呼び出し、結果を作業ウィンドウに返します。

この記事では、Node.js と Express を使用するアドインを使用します。 ASP.NET ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org/) (最新 [LTS](https://nodejs.org/about/releases) バージョン)

- [Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)

- コード エディター - Visual Studio Code をお勧めします

- Microsoft 365 サブスクリプションのOneDrive for Businessに保存されている少なくともいくつかのファイルとフォルダー

- [IdentityAPI 1.3 要件セット](/javascript/api/requirement-sets/common/identity-api-requirement-sets)をサポートする Microsoft 365 のビルド。 更新可能な 90 日間の[Microsoft 365 E5開発者](https://developer.microsoft.com/microsoft-365/dev-program#Subscription)サブスクリプションを提供する無料の開発者サンドボックスを入手できます。 開発者サンドボックスには、この記事の後の手順でアプリの登録に使用できる Microsoft Azure サブスクリプションが含まれています。 必要に応じて、アプリの登録に別の Microsoft Azure サブスクリプションを使用できます。 [Microsoft Azure](https://account.windowsazure.com/SignUp) で試用版サブスクリプションを取得します。

## <a name="set-up-the-starter-project"></a>スタート プロジェクトをセットアップする

1. 「[Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)」にあるリポジトリを複製するかダウンロードします。

   > [!NOTE]
   > サンプルには 2 つのバージョンがあります。
   >
   > - **Begin** フォルダーはスターター プロジェクトです。 SSO や承認に直接関連しない UI などの側面は、既に完了しています。 この記事で後述する各セクションでは、これを完成させるための手順を順に説明します。
   > - **Complete** フォルダーには、この記事のすべてのコーディング手順が完了した同じサンプルが含まれています。 完成したバージョンを使用するには、この記事の手順に従うだけですが、"Begin" を "Complete" に置き換え、「 **クライアント側のコーディング** 」と「 **中間層サーバー側のコーディング** 」のセクションをスキップします。

1. **[開始**] フォルダーでコマンド プロンプトを開きます。

1. コンソールで `npm install` を入力して、package.json ファイルに項目化されているすべての依存関係をインストールします。

1. コマンド`npm run install-dev-certs`を実行します。 証明書をインストールするプロンプトに対して **はい** を選択します。

以降のアプリ登録手順のプレースホルダーには、次の値を使用します。

| プレースホルダー           | 値                                 |
|-----------------------|---------------------------------------|
| `<add-in-name>`       | **Office-Add-in-NodeJS-SSO**          |
| `<redirect-platform>` | **シングルページ アプリケーション (SPA)**     |
| `<redirect-uri>`      | `https://localhost:44355/dialog.html` |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="configure-the-add-in"></a>アドインを構成する

1. コード エディターで複製プロジェクトの`\Begin`フォルダーを開きます。

1. ファイルを `.ENV` 開き、 **Office-Add-in-NodeJS-SSO** アプリの登録から前にコピーした値を使用します。 次のように値を設定します。

   | 名前              | 値                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | アプリ登録の概要ページからの **アプリケーション (クライアント) ID**。 |
   | **CLIENT_SECRET** | **[証明書] & [シークレット**] ページから保存された **クライアント** シークレット。       |
   | **DIRECTORY_ID**  | アプリ登録の概要ページからの **ディレクトリ (テナント) ID**。   |

   値は引用符で囲ま **ない** でください。 完了すると、ファイルは以下のようになります。

   ```javascript
   CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
   CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
   DIRECTORY_ID=478aa78e-20ba-4c0d-9ffe-c4f62e5de3d5
   NODE_ENV=development
   SERVER_SOURCE=https://localhost:44355

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:44355/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. マークアップ内の _両方の場所にある_ プレースホルダー "$app-id-guid$" を、**Office-Add-in-NodeJS-SSO** アプリ登録の作成時にコピーした **アプリケーション ID** に置き換えます。 "$" 記号は ID の一部ではないため、含めないでください。 これは、 のCLIENT_IDに使用した ID と同じです。ENV ファイル。

   > [!NOTE]
   > 値は **\<Resource\>** 、アドインを登録したときに設定した **アプリケーション ID URI** です。 セクションは **\<Scopes\>** 、アドインが AppSource を通じて販売されている場合にのみ、同意ダイアログ ボックスを生成するために使用されます。

1. `\public\javascripts\fallback-msal\authConfig.js` ファイルを開きます。 プレースホルダー "$app-id-guid$" を、前に作成した **Office-Add-in-NodeJS-SSO** アプリ登録から保存したアプリケーション ID に置き換えます。

1. 変更をファイルに保存します。

## <a name="code-the-client-side"></a>クライアント側のコーディング

### <a name="create-client-request-and-response-handler"></a>クライアント要求と応答ハンドラーを作成する

1. コード エディターで、`public\javascripts\ssoAuthES6.js`ファイルを開きます。 Internet Explorer 11 でも Promise がサポートされることを保証するコードと、アドインの唯一のボタンにハンドラーを割り当てるための`Office.onReady`呼び出しが既にあります。

   > [!NOTE]
   > 名前が示すように、ssoAuthES6.js は JavaScript ES6 構文を使用します。これは、これは、`async`と`await`の使用こそが SSO API の本質的なシンプルさを最もよく示すためです。 localhost サーバーが起動されると、このファイルは ES5 構文に変換され、サンプルで Internet Explorer 11 がサポートされます。

    サンプル コードの重要な部分は、クライアント要求です。 クライアント要求は、中間層サーバーで REST API を呼び出すための要求に関する情報を追跡するオブジェクトです。 これは、次のシナリオを通じてクライアント要求の状態を追跡または更新する必要があるため、必要です。

    - SSO が失敗し、フォールバック認証が必要です。 アクセス トークンは、ポップアップ ダイアログ ボックスで MSAL を介して取得されます。 目標は、このシナリオで失敗せず、代替認証アプローチに適切にフォールバックすることです。

    クライアント要求オブジェクトは、次のデータを追跡します。

    - `authSSO` - SSO を使用する場合は true、それ以外の場合は false。
    - `verb` - GET や POST などの REST API 動詞。
    - `accessToken`- ASP.NET Core サーバーへのアクセス トークン。
    - `url`- ASP.NET Core サーバーで呼び出す REST API の URL。
    - `callbackRESTApiHandler` - REST API 呼び出しの結果を渡す関数。
    - `callbackFunction` - 準備ができたときにクライアント要求を渡す関数。

1. クライアント要求オブジェクトを初期化するには、 関数で `createRequest` を次のコードに置き換えます `TODO 1` 。

    ```javascript
    const clientRequest = {
      authSSO: authSSO,
      verb: verb,
      accessToken: null,
      url: url,
      callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };
    ```

1. `TODO 2` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - SSO が使用されているかどうかを確認します。 アクセス トークンを取得する方法は、SSO の場合とフォールバック認証の場合とは異なります。
    - SSO がアクセス トークンを返す場合は、 関数を `callbackfunction` 呼び出します。 フォールバック認証では を呼び出 `dialogFallback`します。これは、最終的にユーザーが MSAL 経由でサインインした後にコールバック関数を呼び出します。

    ```javascript
    // Get access token.

    if (authSSO) {
    try {
      // Get access token from Office SSO.
      clientRequest.accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      callbackFunction(clientRequest);
    } catch (error) {
      // handle the SSO error which will inform us if we need to switch to fallback auth.
      let fallbackRequired = handleSSOErrors(error);
      if (fallbackRequired) switchToFallbackAuth(clientRequest);
    }
   } else {
     // Use fallback auth to get access token.
     dialogFallback(clientRequest);
   }
    ```

1. `getFileNameList` 関数で、`TODO 3` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - この関数 `getFileNameList` は、ユーザーが作業ウィンドウで [ **OneDrive ファイル名の取得** ] ボタンを選択すると呼び出されます。
    - REST API の URL など、呼び出しに関する情報を追跡するクライアント要求が作成されます。
    - REST API が結果を返すと、関数に `handleGetFileNameResponse` 渡されます。 このコールバックは にパラメーター `createRequest` として渡され、 で `clientRequest.callbackRESTApiHandler`追跡されます。
    - このコードは、クライアント要求を呼び出 `callWebServer` して次の手順を実行し、REST API を呼び出します。

    ```javascript
    createRequest(
      "GET",
      "/getuserfilenames",
      handleGetFileNameResponse,
      async (clientRequest) => {
        await callWebServer(clientRequest);
      }
    );
    ```

1. `handleGetFileNameResponse` 関数で、`TODO 4` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - コードは、ドキュメントにファイル名を書き込む応答 (ファイル名の一覧を含む) `writeFileNamesToOfficeDocument` を渡します。
    - コードはエラーをチェックします。 ファイル名が書き込まれている場合は成功メッセージが表示され、それ以外の場合はエラーが表示されます。

    ```javascript
    if (response !== null) {
      try {
        await writeFileNamesToOfficeDocument(response);
        showMessage("Your OneDrive filenames are added to the document.");
      } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin
        // "Unable to add filenames to document."
        showMessage(error);
      }
    } else
    showMessage("A null response was returned to handleGetFileNameResponse.");
    ```

1. `handleSSOErrors` 関数で、`TODO 5` を次のコードに置き換えます。 これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。

    ```javascript
    let fallbackRequired = false;

   switch (err.code) {
     case 13001:
       // No one is signed into Office. If the add-in cannot be effectively used when no one
       // is logged into Office, then the first call of getAccessToken should pass the
       // `allowSignInPrompt: true` option. Since this sample does that, you should not see
       // this error.
       showMessage(
         "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
       );
       break;
     case 13002:
       // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
       // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
       showMessage(
         "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
       );
       break;
     case 13006:
       // Only seen in Office on the web.
       showMessage(
         "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
       );
       break;
     case 13008:
       // Only seen in Office on the web.
       showMessage(
        "Office is still working on the last operation. When it completes, try this operation again."
       );
       break;
     case 13010:
       // Only seen in Office on the web.
       showMessage(
         "Follow the instructions to change your browser's zone configuration."
       );
       break;
    ```

1. `TODO 6`を以下のコードに置き換えます。 これらのエラーの詳細については、「 [Office アドインでの SSO のトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」を参照してください。処理できないエラーについては、 `true` 呼び出し元に返されます。 これは、呼び出し元がフォールバック認証として MSAL を使用するように切り替える必要があることを示します。

    ```javascript
     default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
      // to non-SSO sign-in.
      fallbackRequired = true;
      break;
    }
    return fallbackRequired;
    ```

### <a name="call-the-rest-api-on-the-middle-tier-server"></a>中間層サーバーで REST API を呼び出す

1. `callWebServer` 関数で、`TODO 7` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - 実際の AJAX 呼び出しは、 関数によって `ajaxCallToRESTApi` 行われます。
    - 中間層サーバーから現在のトークンの有効期限が切れたことを示すエラーが返された場合、この関数は新しいアクセス トークンの取得を試みます。
    - AJAX 呼び出しが正常に完了できない場合は、 `switchToFallbackAuth` Office SSO ではなく MSAL 認証を使用するように呼び出されます。

    ```javascript
    try {
    const data = await $.ajax({
      type: clientRequest.verb,
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
    });
    clientRequest.callbackRESTApiHandler(data);

    } catch (error) {
     // TODO 8: Check for expired SSO token and refresh if needed.

    // TODO 9: Check for Microsoft Graph and other errors.

    }
    ```

1. `TODO 8` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - サーバーは、期限切れのトークンを識別すると、"TokenExpiredError" 型のエラーを返します。
    - try...catch は Office.auth.getAccessToken を呼び出して、新しい有効期限で更新されたトークンを取得します。
    - このコードでは、サーバー API の呼び出しが再試行されます。

    ```javascript
    // Check for expired SSO token. Refresh and retry the call if it expired.
    if (
      error.responseJSON &&
      authSSO === true &&
      error.responseJSON.type === "TokenExpiredError"
    ) {
      try {
        const accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
        const data = await $.ajax({
          type: clientRequest.verb,
          url: clientRequest.url,
          headers: { Authorization: "Bearer " + accessToken },
          cache: false,
        });
        clientRequest.callbackRESTApiHandler(data);
      } catch (error) {
        showMessage(error.responseText);
        switchToFallbackAuth(clientRequest);
        return;
      }
    }
    ```

1. `TODO 9` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - **Microsoft Graph** エラーの場合は、作業ウィンドウにメッセージを表示します。
    - その他のすべてのメッセージの場合は、作業ウィンドウにメッセージを表示します。

    ```javascript
    // Check for a Microsoft Graph API call error. which is returned as bad request (403)
    if (error.status === 403) {
      if (error.responseJSON && error.responseJSON.type === "Microsoft Graph") {
        showMessage(error.responseJSON.errorDetails);
      } else {
        showMessage(error);
      }
      return;
    }

    // For all other error scenarios, display the message and use fallback auth.
    showMessage("Unknown error from web server: " + JSON.stringify(error));
    if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    ```

フォールバック認証では、MSAL ライブラリを使用してユーザーにサインインします。 アドイン自体は SPA であり、SPA アプリの登録を使用して中間層サーバーにアクセスします。

1. `switchToFallbackAuth` 関数で、`TODO 10` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - グローバル `authSSO` を false に設定し、認証に MSAL を使用する新しいクライアント要求を作成します。新しい要求には、中間層サーバーへの MSAL アクセス トークンがあります。
    - 要求が作成されたら、 を呼び出 `callWebServer` して、引き続き中間層サーバーを正常に呼び出そうとします。

    ```javascript
    // Guard against accidental call to this function when fallback is already in use.

    if (authSSO === false) return;

    showMessage("Switching from SSO to fallback auth.");
    authSSO = false;
    // Create a new request for fallback auth.
    createRequest(
      clientRequest.verb,
      clientRequest.url,
      clientRequest.callbackRESTApiHandler,
      async (fallbackRequest) => {
        // Hand off to call using fallback auth.
        await callWebServer(fallbackRequest);
      }
    );
    ```

## <a name="code-the-middle-tier-server"></a>中間層サーバーをコーディングする

中間層サーバーは、クライアントが呼び出す REST API を提供します。 たとえば、REST API `/getuserfilenames` は、ユーザーの OneDrive フォルダーからファイル名の一覧を取得します。 各 REST API 呼び出しでは、正しいクライアントがデータにアクセスしていることを確認するために、クライアントによるアクセス トークンが必要です。 アクセス トークンは、On-Behalf-Of フロー (OBO) を介して Microsoft Graph トークンと交換されます。 新しい Microsoft Graph トークンは、後続の API 呼び出しのために MSAL ライブラリによってキャッシュされます。 中間層サーバーの外部に送信されることはありません。 詳細については、「[中間層のアクセス トークン要求](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)」を参照してください。

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>ルートを作成し、On-Behalf-Of フローを実装する

1. ファイル `routes\getFilesRoute.js` を開き、次のコードに置き換えます `TODO 11` 。 このコードについては、以下の点に注意してください。

    - を呼び出します `authHelper.validateJwt`。 これにより、アクセス トークンが有効であり、改ざんされていないことが保証されます。
    - 詳細については、「 [トークンの検証](/azure/active-directory/develop/access-tokens#validating-tokens)」を参照してください。

    ```javascript
    router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 12: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
    );
    ```

1. `TODO 12` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - 必要な最小スコープ (など `files.read`) のみが要求されます。
    - MSAL `authHelper` を使用して への呼び出し `acquireTokenOnBehalfOf`で OBO フローを実行します。

    ```javascript
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(" ")[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
      const accessAsUserScope = tokenScopes.find(
        (scope) => scope === "access_as_user"
      );
      if (!accessAsUserScope) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // TODO 13: Call Microsoft Graph to get list of filenames.
    } catch (err) {
      // TODO 14: Handle any errors.
    }
    ```

1. `TODO 13` を次のコードに置き換えます。 このコードについては、以下の点に注意してください。

    - Microsoft Graph API 呼び出しの URL を構築し、関数を介して呼び出しを`getGraphData`行います。
    - HTTP 500 応答と詳細を送信してエラーを返します。
    - 成功すると、ファイル名リストを含む JSON がクライアントに返されます。

    ```javascript
    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = "/me/drive/root/children";

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = "?$select=name&$top=10";

    const graphData = await getGraphData(
      response.accessToken,
      rootUrl,
      params
    );

    // If Microsoft Graph returns an error, such as invalid or expired token,
    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
    // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
    if (graphData.code) {
      res
        .status(403)
        .send({
          type: "Microsoft Graph",
          errorDetails:
            "An error occurred while calling the Microsoft Graph API.\n" +
            graphData,
        });
    } else {
      // MS Graph data includes OData metadata and eTags that we don't need.
      // Send only what is actually needed to the client: the item names.
      const itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
      }

      res.status(200).send(itemNames);
    }
    ```

1. `TODO 14` を次のコードに置き換えます。 このコードでは、クライアントが新しいトークンを要求して再度呼び出すことができるため、トークンの有効期限が切れたかどうかを特に確認します。

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
   } else {
     res.status(403).send({ type: "Unknown", errorDetails: err });
   }
   ```

このサンプルでは、MSAL によるフォールバック認証と Office 経由の SSO 認証の両方を処理する必要があります。 サンプルでは最初に SSO を試し、サンプルが SSO を `authSSO` 使用している場合、またはフォールバック認証に切り替えた場合は、ファイルの上部にあるブール値が追跡されます。

## <a name="run-the-project"></a>プロジェクトを実行する

1. 結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。

1. `\Begin`フォルダーのルートでコマンド プロンプトを開きます。

1. コマンドを実行して、すべてのパッケージの依存関係をインストールします `npm install` 。

1. コマンド `npm start` を実行して中間層サーバーを起動します。

1. アドインを Office アプリケーション (Excel、Word、または PowerPoint) にサイドロードして、テストをする必要があります。 手順はプラットフォームによって異なります。 「[テスト用に Office アドインをサイドロードする](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)」に手順へのリンクがあります。

1. Office アプリケーションの **[ホーム]** リボンで **[アドインの表示]** ボタン (**SSO Node.js** グループ内) を選択して、作業ウィンドウ アドインを開きます。

1. **[OneDrive ファイル名の取得]** ボタンをクリックします。 Microsoft 365 Educationまたは職場アカウント、または Microsoft アカウントを使用して Office にログインしていて、SSO が正常に動作している場合は、OneDrive for Business内の最初の 10 個のファイルとフォルダー名がドキュメントに挿入されます。 (初回は 15 秒ほどかかる場合があります。ログインしていない場合、または SSO をサポートしていないシナリオや、何らかの理由で SSO が機能しない場合は、サインインするように求められます。 サインインすると、ファイル名とフォルダー名が表示されます。

> [!NOTE]
> 以前に別の ID で Office にサインインしており、その時点で開いていた一部の Office アプリケーションがまだ開いている場合、Office が ID を変更したかのように見えても、確実に ID を変更できていない場合があります。 これが発生すると、Microsoft Graph の呼び出しが失敗するか、以前の ID のデータが返される場合があります。 これを防ぐには、必ず _他のすべての Office アプリケーションを閉じて_ から、**[OneDrive ファイル名の取得]** を押してください。

## <a name="security-notes"></a>セキュリティに関する注意事項

- の`getFilesroute.js`ルートでは`/getuserfilenames`、リテラル文字列を使用して Microsoft Graph の呼び出しを作成します。 文字列の任意の部分がユーザー入力から取得されるように呼び出しを変更する場合は、応答ヘッダーインジェクション攻撃で使用できないように入力をサニタイズします。

- 次のコンテンツ では `app.js` 、スクリプトに対してセキュリティ ポリシーが適用されます。 アドインのセキュリティニーズに応じて、追加の制限を指定することもできます。

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

[Microsoft ID プラットフォームドキュメント](/azure/active-directory/develop/)のセキュリティのベスト プラクティスに常に従ってください。
